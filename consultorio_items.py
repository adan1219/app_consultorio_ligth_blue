#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Backend: consultorio_items.py
- Soporta modelo normalizado: Consultas (cabecera) + Consulta_Items (líneas)
- Lee/writes Excel definido en config.json
- Servicios: RegistroService (crear consulta + items), ReporteService (filtrado por items),
  CorteService (calcula comisiones por item) y KPIService (usa items para top tratamientos).
- Migración automática: si existen filas en Consultas con tratamiento y costo_servicio,
  se crea un item por consulta al inicializar (no borra nada).
"""

from __future__ import annotations
import argparse
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import sys
from datetime import datetime, date
from dateutil import parser as dtparser
import pandas as pd
import numpy as np
import unicodedata
from difflib import get_close_matches

# ---------------------------
# Config
# ---------------------------

@dataclass
class AppConfig:
    excel_path: Path
    overwrite_template: bool = False
    promoter_percent: float = 0.15
    doctors_percent: Dict[str, float] = field(default_factory=dict)
    pdf_output_dir: Optional[Path] = None
    admin_password_hash: Optional[str] = None  # hash hex string (SHA256) para proteger la pestaña Config

    @classmethod
    def from_json(cls, config_path: Path) -> "AppConfig":
        raw = json.loads(config_path.read_text(encoding="utf-8"))
        excel_path = Path(raw.get("excel_path", "./consultorio.xlsx")).expanduser().resolve()
        pdf_dir = raw.get("pdf_output_dir")
        return cls(
            excel_path=excel_path,
            overwrite_template=bool(raw.get("overwrite_template", False)),
            promoter_percent=float(raw.get("promoter_percent", 0.15)),
            doctors_percent=dict(raw.get("doctors_percent", {})),
            pdf_output_dir=(Path(pdf_dir).expanduser().resolve() if pdf_dir else excel_path.parent),
            admin_password_hash=raw.get("admin_password_hash")  # puede ser None
        )


# ---------------------------
# Schema Excel
# ---------------------------

SHEETS = {
    "Pacientes": ["id_paciente", "nombre", "telefono", "fecha_alta", "folio_paciente"],
    # Personal ahora incluye pct_comision y metadatos
    "Personal": ["id_personal", "nombre", "rol", "activo", "pct_comision", "creado_por", "creado_en", "modificado_por", "modificado_en"],
    "Tratamientos": ["id_tratamiento", "nombre", "activo"],
    "Consultas": ["id_consulta", "id_paciente", "fecha", "id_doctor", "id_promotor", "estado", "total_consulta", "metodo_pago", "moneda", "notas"],
    "Consulta_Items": [
    "id_item", "id_consulta","id_tratamiento" ,"descripcion_item",
    "cantidad", "precio_unitario", "subtotal",
    "id_tecnico_item"],
    "Cortes": ["id_corte", "fecha_corte", "rango_inicio", "rango_fin",
              "total_ingresos", "total_costo_tecnico", "total_comisiones_doctores", "total_comisiones_promotores", "total_utilidad", "creado_por", "creado_en", "version"],
    "Corte_Detalle": ["id_corte_detalle", "id_corte", "id_consulta", "id_item", "doctor", "promotor", "tecnico", "comision_doctor", "comision_promotor", "costo_tecnico", "incluido"],
    "KPIs_Export": ["fecha_inicio", "fecha_fin", "ingresos_periodo", "pacientes_unicos", "ticket_promedio", "top_tratamientos_json", "top_dentistas_json", "retencion", "utilidad_periodo"],
    # historial de cambios al personal
    "Personal_History": ["fecha", "usuario", "accion", "id_personal", "nombre", "rol", "valor_anterior", "valor_nuevo"],
    "Citas": ["id_cita", "id_paciente","fecha", "hora", "hora_fin","id_doctor", "estado", "notas"],
     "Comisiones_Abonos": ["id_abono", "id_personal", "id_consulta", "fecha_pago", "monto", "moneda", "nota"]
}




# ---------------------------
# Utilitarios
# ---------------------------

def _ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

def _new_id(prefix: str, series: pd.Series) -> str:
    nums = []
    for v in series.dropna().astype(str):
        try:
            nums.append(int("".join([ch for ch in v if ch.isdigit()])))
        except:
            continue
    nxt = (max(nums) + 1) if nums else 1
    return f"{prefix}{nxt:04d}"

def _parse_date(s: str) -> date:
    return dtparser.parse(s).date()

def normalize_name(s: Optional[str]) -> str:
    """Normaliza texto: lower, elimina acentos y espacios redundantes."""
    if not s:
        return ""
    s = str(s).strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return " ".join(s.split())

# coerción de tipos al leer/escribir para evitar merges fallidos
def _coerce_id_like_columns(dfs: Dict[str, pd.DataFrame]) -> None:
    id_map = {
        "Pacientes": ["id_paciente"],
        "Personal": ["id_personal"],
        "Tratamientos": ["id_tratamiento"],
        "Consultas": ["id_consulta", "id_paciente", "id_doctor", "id_promotor"],
        "Consulta_Items": ["id_item","id_consulta","id_tratamiento","id_tecnico_item"],
        "Corte_Detalle": ["id_corte_detalle","id_corte","id_consulta","id_item"],
        "Comisiones_Abonos": ["id_abono","id_personal","id_consulta"],
    }
    for sheet, cols in id_map.items():
        if sheet in dfs:
            df = dfs[sheet]
            for c in cols:
                if c in df.columns:
                    df[c] = df[c].astype("string").fillna("").str.replace(r"\.0$", "", regex=True)
    # fechas
    if "Consultas" in dfs and "fecha" in dfs["Consultas"].columns:
        dfs["Consultas"]["fecha"] = pd.to_datetime(dfs["Consultas"]["fecha"], errors="coerce")
    if "Cortes" in dfs and "fecha_corte" in dfs["Cortes"].columns:
        dfs["Cortes"]["fecha_corte"] = pd.to_datetime(dfs["Cortes"]["fecha_corte"], errors="coerce")

# ---------------------------
# I/O Excel
# ---------------------------

def _init_template(path: Path, overwrite: bool) -> None:
    if path.exists() and not overwrite:
        raise FileExistsError(f"Archivo ya existe: {path}")
    _ensure_parent(path)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for sheet, cols in SHEETS.items():
            pd.DataFrame(columns=cols).to_excel(xw, sheet_name=sheet, index=False)

def _read_excel(path: Path) -> Dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(f"No existe el Excel: {path}")
    xls = pd.ExcelFile(path)
    
    dfs = {}

    for name, cols in SHEETS.items():
        if name in xls.sheet_names:
            dfs[name] = xls.parse(name)
        else:
            # Crear hoja vacía SOLO SI NO EXISTE EN EL ARCHIVO
            # pero NO generar una hoja vacía cuando ya existía en memoria
            dfs[name] = pd.DataFrame(columns=cols)


    _coerce_id_like_columns(dfs)
    return dfs

def _write_excel(path: Path, dfs: Dict[str, pd.DataFrame]) -> None:
    # Asegura columnas and tipos mínimos
    _coerce_id_like_columns(dfs)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for sheet, cols in SHEETS.items():
            df = dfs.get(sheet, pd.DataFrame(columns=cols))
            for c in cols:
                if c not in df.columns:
                    df[c] = pd.Series(dtype="object")
            df = df[cols]
            df.to_excel(xw, sheet_name=sheet, index=False)

# ---------------------------
# Repo (CRUD simplificado)
# ---------------------------

class Repo:
    def __init__(self, cfg: AppConfig):
        self.cfg = cfg
        self.dfs = _read_excel(cfg.excel_path)
        print("\n==== DEBUG 1: Consulta_Items CRUDO ====")
        print(self.dfs["Consulta_Items"].head(20))

        # === Inicializar historial de personal ===
        # Trabajar SIEMPRE con Personal_History
        if "Personal_History" in self.dfs:
            dfh = self.dfs["Personal_History"].copy()
            needed = ["fecha","usuario","accion","id_personal","nombre","rol","valor_anterior","valor_nuevo"]
            for c in needed:
                if c not in dfh.columns:
                    dfh[c] = ""
            self.df_hist_personal = dfh.copy()
        else:
            # crear hoja nueva vacía
            self.df_hist_personal = pd.DataFrame(columns=[
                "fecha","usuario","accion","id_personal","nombre","rol","valor_anterior","valor_nuevo"
            ])

        # 👈 CLAVE: Mantener sincronizado el dict de hojas
        self.dfs["Personal_History"] = self.df_hist_personal.copy()

        # resto igual...
        if "Citas" not in self.dfs:
            self.dfs["Citas"] = pd.DataFrame(columns=SHEETS["Citas"])

        self._auto_migrate_consultas_to_items()

        if "folio_paciente" not in self.dfs["Pacientes"].columns:
            self.dfs["Pacientes"]["folio_paciente"] = ""

        self._ensure_consulta_items_columns()


            

    def save(self):
        _write_excel(self.cfg.excel_path, self.dfs)

    def _auto_migrate_consultas_to_items(self):
        # Si Consulta_Items está vacía y hay Consultas con tratamiento/costo antiguo
        ci = self.dfs["Consulta_Items"]
        if not ci.empty:
            return
        # Checa si en Consultas hay columna 'id_tratamiento' o 'costo_servicio' (modelo viejo)
        old_cons = self.dfs["Consultas"]
        # Modelo antiguo: existía columna 'id_tratamiento' o 'tratamiento' o 'costo_servicio' en Consultas
        possible_cols = ["id_tratamiento","tratamiento","costo_servicio"]
        if not any(c in old_cons.columns for c in possible_cols):
            return
        # Intentamos migrar: por cada fila de Consultas, crear 1 item si tiene 'tratamiento' y 'total_consulta' o 'costo_servicio'
        migrated = []
        for _, r in old_cons.iterrows():
            cid = str(r.get("id_consulta","")).strip()
            if not cid:
                continue
            # determinar tratamiento
            tid = r.get("id_tratamiento", "")
            desc = r.get("tratamiento", "")
            price = r.get("costo_servicio", r.get("total_consulta", np.nan))
            if pd.isna(price) or price == "":
                continue
            # crear item
            new_item = {
                "id_item": _new_id("I", ci["id_item"]) if not ci.empty else _new_id("I", pd.Series([])),
                "id_consulta": cid,
                "id_tratamiento": tid if tid else "",
                "descripcion_item": desc if desc and not pd.isna(desc) else "",
                "cantidad": 1,
                "precio_unitario": float(price),
                "subtotal": float(price),
                "id_tecnico_item": "",
                "observaciones_item": ""
            }
            migrated.append(new_item)
            # append into ci
            if ci.empty:
                ci = pd.DataFrame([new_item])
            else:
                ci = pd.concat([ci, pd.DataFrame([new_item])], ignore_index=True)
        self.dfs["Consulta_Items"] = ci
        if len(migrated) > 0:
            self.save()

    def _ensure_consulta_items_columns(self):
        """
        Garantiza que la hoja Consulta_Items tenga todas las columnas requeridas
        sin borrar información existente.
        """
        df = self.dfs.get("Consulta_Items", pd.DataFrame())
        if df.empty:
            # crear estructura vacía
            df = pd.DataFrame(columns=[
                "id_item", "id_consulta", "descripcion_item",
                "cantidad", "precio_unitario", "subtotal", "id_tecnico_item"
            ])
        else:
            # agregar las columnas que falten
            required = ["id_item", "id_consulta", "id_tratamiento", "descripcion_item",
                        "cantidad", "precio_unitario", "subtotal", "id_tecnico_item"]
            for c in required:
                if c not in df.columns:
                    if c in ("cantidad", "precio_unitario", "subtotal"):
                        df[c] = 0.0
                    else:
                        df[c] = ""
            # convertir numéricos
            df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(1.0)
            df["precio_unitario"] = pd.to_numeric(df["precio_unitario"], errors="coerce").fillna(0.0)
            df["subtotal"] = pd.to_numeric(df["subtotal"], errors="coerce").fillna(
                df["cantidad"] * df["precio_unitario"]
            )
        self.dfs["Consulta_Items"] = df


    # Pacientes
    def find_patient_by_name(self, name: str) -> pd.DataFrame:
        n = (name or "").strip().lower()
        df = self.dfs["Pacientes"]
        return df[df["nombre"].fillna("").str.lower().str.contains(n, na=False)]

    def find_patient_by_phone(self, phone: str) -> pd.DataFrame:
        p = (phone or "").strip()
        df = self.dfs["Pacientes"]
        return df[df["telefono"].astype(str).str.contains(p, na=False)]

    def upsert_patient(self, nombre: str, telefono: Optional[str]) -> Tuple[str, Optional[str], bool]:
        df = self.dfs["Pacientes"]
        telefono = (telefono or "").strip()
        mask = (df["nombre"].fillna("").str.lower() == nombre.strip().lower())
        if telefono:
            mask |= (df["telefono"].astype(str) == telefono)
        hit = df[mask]
        if not hit.empty:
            row = hit.iloc[0]
            pid = row["id_paciente"]
            if telefono and (not row.get("telefono") or str(row["telefono"]).strip() == ""):
                self.dfs["Pacientes"].loc[self.dfs["Pacientes"]["id_paciente"] == pid, "telefono"] = telefono
            return pid, telefono, True
        pid = _new_id("P", df["id_paciente"])
        now = datetime.now().date().isoformat()
        self.dfs["Pacientes"] = pd.concat([df, pd.DataFrame([{"id_paciente": pid, "nombre": nombre, "telefono": telefono, "fecha_alta": now}])], ignore_index=True)
        return pid, telefono, False

    def update_paciente_folio(self, id_paciente: str, folio: str):
        """Actualiza o asigna el número de folio clínico al paciente."""
        df = self.dfs["Pacientes"]
        idx = df.index[df["id_paciente"] == id_paciente].tolist()
        if not idx:
            return False
        i = idx[0]
        df.at[i, "folio_paciente"] = folio.strip()
        self.dfs["Pacientes"] = df
        self.save()
        return True

     # ---------------------------------------------------
    # Citas 
    # ---------------------------------------------------
    def crear_cita(self, id_paciente, fecha, hora, id_doctor="", notas="", hora_fin=None):
        df = self.dfs["Citas"]

        # Crear ID único
        id_cita = f"C{datetime.now().strftime('%y%m%d%H%M%S')}"

        nuevo = {
            "id_cita": id_cita,
            "id_paciente": id_paciente,
            "fecha": fecha,
            "hora": hora,
            "hora_fin": hora_fin or "",  # si no viene, queda vacío
            "id_doctor": id_doctor,
            "estado": "pendiente",
            "notas": notas,
        }

        self.dfs["Citas"] = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)
        self.save()
        return id_cita

    def citas_por_paciente(self, id_paciente, solo_pendientes=True):
        df = self.dfs["Citas"]
        if df.empty:
            return pd.DataFrame()
        sub = df[df["id_paciente"] == id_paciente].copy()
        if solo_pendientes:
            sub = sub[sub["estado"].isin(["programada","reprogramada"])]
        sub["fecha"] = pd.to_datetime(sub["fecha"], errors="coerce").dt.date
        return sub.sort_values(["fecha","hora"])

    def citas_en_rango(self, fecha_ini, fecha_fin):
        df = self.dfs["Citas"]
        if df.empty:
            return df.copy()

        # Normalizar: convertir fecha del DataFrame a string YYYY-MM-DD
        df2 = df.copy()
        df2["fecha"] = df2["fecha"].astype(str)

        # Normalizar fecha_ini y fecha_fin
        if not isinstance(fecha_ini, str):
            fecha_ini = fecha_ini.isoformat()
        if not isinstance(fecha_fin, str):
            fecha_fin = fecha_fin.isoformat()

        df2 = df2[(df2["fecha"] >= fecha_ini) & (df2["fecha"] <= fecha_fin)]
        return df2.copy()


    def actualizar_estado_cita(self, id_cita, nuevo_estado):
        df = self.dfs["Citas"]
        idx = df.index[df["id_cita"] == id_cita].tolist()
        if not idx:
            return False
        self.dfs["Citas"].at[idx[0], "estado"] = nuevo_estado
        self.save()
        return True

    # Tratamiento
    def ensure_tratamiento(self, nombre: str) -> str:
        df = self.dfs["Tratamientos"]
        row = df[df["nombre"].fillna("").str.lower() == (nombre or "").strip().lower()]
        if not row.empty:
            return row.iloc[0]["id_tratamiento"]
        tid = _new_id("T", df["id_tratamiento"])
        self.dfs["Tratamientos"] = pd.concat([df, pd.DataFrame([{"id_tratamiento": tid, "nombre": nombre, "activo": 1}])], ignore_index=True)
        return tid

    # Personal
    def ensure_personal(self, nombre: Optional[str], rol: str) -> Optional[str]:
        if not nombre:
            return None
        df = self.dfs["Personal"]
        row = df[(df["nombre"].fillna("").str.lower() == (nombre or "").strip().lower()) & (df["rol"].str.lower() == rol)]
        if not row.empty:
            return row.iloc[0]["id_personal"]
        pid = _new_id("R", df["id_personal"])
        self.dfs["Personal"] = pd.concat([df, pd.DataFrame([{"id_personal": pid, "nombre": nombre, "rol": rol, "activo": 1}])], ignore_index=True)
        return pid

    def find_personal_by_name(self, nombre: str, rol: Optional[str] = None) -> pd.DataFrame:
        """
        Busca filas en Personal por nombre (case-insensitive, con normalización).
        Si se da rol, también filtra por rol.
        """
        df = self.dfs.get("Personal", pd.DataFrame()).copy()
        if df.empty:
            return df

        df["nombre_norm"] = df["nombre"].fillna("").apply(normalize_name)
        target = normalize_name(nombre)
        sub = df[df["nombre_norm"] == target]
        if rol:
            sub = sub[sub["rol"].astype(str).str.lower() == rol.strip().lower()]
        return sub

########################################################################################################################################################

        # --- Métodos para administrar personal con porcentaje y llevar historial ---
    def list_personal(self) -> pd.DataFrame:
        """Devuelve copia del DataFrame Personal con columnas aseguradas."""
        df = self.dfs.get("Personal", pd.DataFrame()).copy()
        # asegurar columnas nuevas
        for c in ["pct_comision", "creado_por", "creado_en", "modificado_por", "modificado_en"]:
            if c not in df.columns:
                df[c] = ""
        return df

    def add_personal(self, nombre: str, rol: str, pct: float, usuario: str = "admin") -> str:
        """Agrega una fila en Personal y registra en Personal_History. Devuelve id_personal."""
        df = self.dfs.get("Personal", pd.DataFrame()).copy()
        pid = _new_id("R", df["id_personal"]) if not df.empty else _new_id("R", pd.Series([]))
        now = datetime.now().isoformat(timespec="seconds")
        nueva = {"id_personal": pid, "nombre": nombre, "rol": rol, "activo": 1, "pct_comision": float(pct), "creado_por": usuario, "creado_en": now, "modificado_por": usuario, "modificado_en": now}
        self.dfs["Personal"] = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
        # historial
        hist = self.dfs.get("Personal_History", pd.DataFrame(columns=SHEETS["Personal_History"]))
        hist_row = {"fecha": now, "usuario": usuario, "accion": "alta_personal", "id_personal": pid, "nombre": nombre, "rol": rol, "valor_anterior": "", "valor_nuevo": str(pct)}
        self.dfs["Personal_History"] = pd.concat([hist, pd.DataFrame([hist_row])], ignore_index=True)
        self.save()
        return pid

    def update_personal_pct(self, id_personal: str, new_pct: float, usuario: str = "admin") -> bool:
        """Actualiza pct_comision de un personal y deja registro en Personal_History."""
        df = self.dfs.get("Personal", pd.DataFrame())
        if df.empty or id_personal not in df["id_personal"].astype(str).tolist():
            return False
        idx = df.index[df["id_personal"].astype(str) == str(id_personal)].tolist()
        if not idx:
            return False
        i = idx[0]
        old_val = df.at[i, "pct_comision"] if "pct_comision" in df.columns else ""
        now = datetime.now().isoformat(timespec="seconds")
        self.dfs["Personal"].at[i, "pct_comision"] = float(new_pct)
        self.dfs["Personal"].at[i, "modificado_por"] = usuario
        self.dfs["Personal"].at[i, "modificado_en"] = now
        # historial
        hist = self.dfs.get("Personal_History", pd.DataFrame(columns=SHEETS["Personal_History"]))
        hist_row = {"fecha": now, "usuario": usuario, "accion": "cambio_pct", "id_personal": id_personal, "nombre": df.at[i,"nombre"], "rol": df.at[i,"rol"], "valor_anterior": str(old_val), "valor_nuevo": str(new_pct)}
        self.dfs["Personal_History"] = pd.concat([hist, pd.DataFrame([hist_row])], ignore_index=True)
        self.save()
        return True

    def deactivate_personal(self, id_personal: str, usuario: str = "admin") -> bool:
        df = self.dfs.get("Personal", pd.DataFrame())
        if df.empty or id_personal not in df["id_personal"].astype(str).tolist():
            return False
        idx = df.index[df["id_personal"].astype(str) == str(id_personal)].tolist()[0]
        now = datetime.now().isoformat(timespec="seconds")
        self.dfs["Personal"].at[idx, "activo"] = 0
        self.dfs["Personal"].at[idx, "modificado_por"] = usuario
        self.dfs["Personal"].at[idx, "modificado_en"] = now
        hist = self.dfs.get("Personal_History", pd.DataFrame(columns=SHEETS["Personal_History"]))
        pr = self.dfs["Personal"].loc[self.dfs["Personal"]["id_personal"] == id_personal].iloc[0]
        hist_row = {"fecha": now, "usuario": usuario, "accion": "baja_personal", "id_personal": id_personal, "nombre": pr.get("nombre",""), "rol": pr.get("rol",""), "valor_anterior": "", "valor_nuevo": ""}
        self.dfs["Personal_History"] = pd.concat([hist, pd.DataFrame([hist_row])], ignore_index=True)
        self.save()
        return True

    def eliminar_personal(self, id_personal: str, usuario: str = "admin") -> bool:
        """
        Elimina definitivamente un registro de Personal y deja traza en Personal_History
        con nombre y rol, para que se vea en el historial.
        """
        df = self.dfs.get("Personal", pd.DataFrame())
        if df.empty:
            return False

        mask = df["id_personal"].astype(str) == str(id_personal)
        if not mask.any():
            return False

        fila = df[mask].iloc[0]
        nombre = fila.get("nombre", "")
        rol = fila.get("rol", "")

        now = datetime.now().isoformat(timespec="seconds")

        # Historial
        hist = self.dfs.get("Personal_History", pd.DataFrame(columns=SHEETS["Personal_History"]))
        hist_row = {
            "fecha": now,
            "usuario": usuario,
            "accion": "eliminado",
            "id_personal": id_personal,
            "nombre": nombre,
            "rol": rol,
            "valor_anterior": "",
            "valor_nuevo": ""
        }
        self.dfs["Personal_History"] = pd.concat([hist, pd.DataFrame([hist_row])], ignore_index=True)

        # Eliminar del maestro de Personal
        self.dfs["Personal"] = df[~mask].reset_index(drop=True)

        self.save()
        return True

    def get_personal_history(self, limit: int = 200) -> pd.DataFrame:
        """
        Devuelve el historial real desde la hoja Personal_History.
        Asegura columnas y elimina NaN.
        """
        df = self.dfs.get("Personal_History")

        # Si no existe la hoja, crearla vacía correctamente
        if df is None:
            df = pd.DataFrame(columns=SHEETS["Personal_History"])
            self.dfs["Personal_History"] = df
            self.save()
            return df

        if df.empty:
            return df

        # Asegurar columnas correctas
        for col in SHEETS["Personal_History"]:
            if col not in df.columns:
                df[col] = ""

        # Quitar NaN
        df = df.fillna("")

        # Ordenar por fecha
        df = df.sort_values("fecha", ascending=False).reset_index(drop=True)

        return df.head(limit)



##################################  cambios por revision de flujo  ##################################


    def add_consulta_en_espera(self, paciente: str, telefono: str, doctor: Optional[str], promotor: Optional[str]) -> str:
        """Crea una nueva consulta en estado 'en espera'."""
        pid, _, _ = self.upsert_patient(paciente, telefono)
        id_doc = self.ensure_personal(doctor, "doctor") if doctor else ""
        id_prom = self.ensure_personal(promotor, "promotor") if promotor else ""
        cons_df = self.dfs["Consultas"]
        cid = f"C{datetime.now().strftime('%y%m%d%H%M%S')}"
        nueva = {
            "id_consulta": cid,
            "id_paciente": pid,
            "fecha": datetime.now(),
            "id_doctor": id_doc,
            "id_promotor": id_prom,
            "estado": "en espera",
            "total_consulta": 0.0,
            "metodo_pago": "",
            "moneda": "",
            "notas": ""
        }
        self.dfs["Consultas"] = pd.concat([cons_df, pd.DataFrame([nueva])], ignore_index=True)
        self.save()
        return cid

##########################################################################################################################################################
    # Crear consulta (cabecera) y items
    def add_consulta_with_items(self, id_paciente: str, fecha: datetime, id_doctor: Optional[str], id_promotor: Optional[str], items: List[Dict], estado: str = "vigente") -> Tuple[str, List[str]]:
        # items: list of { "id_tratamiento", "descripcion_item", "cantidad", "precio_unitario", "id_tecnico_item", "observaciones_item" }
        cons_df = self.dfs["Consultas"]
        cid = _new_id("C", cons_df["id_consulta"])
        total = sum((float(it.get("cantidad",1)) * float(it.get("precio_unitario",0))) for it in items)
        nueva = {"id_consulta": cid, "id_paciente": id_paciente, "fecha": fecha, "id_doctor": id_doctor or "", "id_promotor": id_promotor or "", "estado": estado, "total_consulta": float(total)}
        self.dfs["Consultas"] = pd.concat([cons_df, pd.DataFrame([nueva])], ignore_index=True)

        # items
        ci = self.dfs["Consulta_Items"]
        created_ids = []
        for it in items:
            iid = _new_id("I", ci["id_item"]) if not ci.empty else _new_id("I", pd.Series([]))
            subtotal = float(it.get("cantidad",1)) * float(it.get("precio_unitario",0))
            row = {
                "id_item": iid,
                "id_consulta": cid,
                "id_tratamiento": it.get("id_tratamiento",""),
                "descripcion_item": it.get("descripcion_item", ""),
                "cantidad": float(it.get("cantidad",1)),
                "precio_unitario": float(it.get("precio_unitario",0)),
                "subtotal": float(subtotal),
                "id_tecnico_item": it.get("id_tecnico_item",""),
                "observaciones_item": it.get("observaciones_item","")
            }
            if ci.empty:
                ci = pd.DataFrame([row])
            else:
                ci = pd.concat([ci, pd.DataFrame([row])], ignore_index=True)
            created_ids.append(iid)
        self.dfs["Consulta_Items"] = ci
        self.save()
        return cid, created_ids

    # Lecturas para reportes
    def get_consultas_items_join(self) -> pd.DataFrame:
        # join consultas + items + pacientes + personal + tratamientos
        cons = self.dfs["Consultas"].copy()
        items = self.dfs["Consulta_Items"].copy()
        pac = self.dfs["Pacientes"][["id_paciente","nombre","telefono"]].rename(columns={"nombre":"paciente","telefono":"telefono_paciente"})
        per = self.dfs["Personal"][["id_personal","nombre","rol"]]
        trat = self.dfs["Tratamientos"][["id_tratamiento","nombre"]].rename(columns={"nombre":"tratamiento"})
        # merges
        df = items.merge(cons, on="id_consulta", how="left", suffixes=("_item","_cons"))

        



        df = df.merge(pac, left_on="id_paciente", right_on="id_paciente", how="left")
        # join doctor/promotor/tecnico names from Personal via id's in Consultas (cabecera) or id_tecnico_item
        per_doc = per.add_prefix("doc_")
        per_pro = per.add_prefix("pro_")
        per_tec = per.add_prefix("tec_")
        df = df.merge(per_doc, left_on="id_doctor", right_on="doc_id_personal", how="left")
        df = df.merge(per_pro, left_on="id_promotor", right_on="pro_id_personal", how="left")
        df = df.merge(per_tec, left_on="id_tecnico_item", right_on="tec_id_personal", how="left")
        df = df.merge(trat, on="id_tratamiento", how="left")

        print("\n==== DEBUG 2B: Después merge con TRATAMIENTOS ====")
        print(df[["id_item","descripcion_item","id_tratamiento","tratamiento"]].head(20))


        # normalize numeric
        df["subtotal"] = pd.to_numeric(df["subtotal"], errors="coerce").fillna(0.0)
        df["precio_unitario"] = pd.to_numeric(df["precio_unitario"], errors="coerce").fillna(0.0)
        df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(1.0)
        # friendly columns
        df = df.rename(columns={
            "doc_nombre":"doc_nombre",
            "pro_nombre":"pro_nombre",
            "tec_nombre":"tec_nombre"
        })
        # Asegurar columnas
        for c in ["doc_nombre","pro_nombre","tec_nombre","fecha","paciente","tratamiento",
                "subtotal","precio_unitario","cantidad","id_consulta","id_item"]:
            if c not in df.columns:
                df[c] = ""

        # Tratamiento final
        df["tratamiento"] = df["tratamiento"].fillna("")
        df["tratamiento"] = df.apply(
            lambda r: r["descripcion_item"] if r["tratamiento"] == "" else r["tratamiento"],
            axis=1
        )


        print("\n==== DEBUG 2C: Después de regla final TRATAMIENTO ====")
        print(df[["id_item","descripcion_item","tratamiento"]].head(20))
        return df


    def save_historial_personal(self):
        """
        Guarda el historial unificado en Personal_Historial.
        """
        self.dfs["Personal_History"] = self.df_hist_personal.copy()
        self.save()
# Servicios
# ---------------------------
    def get_notas_paciente(self, id_paciente):
        """
        Devuelve un DataFrame con el historial de notas del paciente
        ordenado por fecha descendente.
        """
        df = self.dfs.get("Consultas", pd.DataFrame()).copy()

        if df.empty:
            return pd.DataFrame(columns=["fecha", "notas"])

        df = df[df["id_paciente"] == id_paciente]

        # solo consultas con notas no vacías
        df = df[df["notas"].astype(str).str.strip() != ""]

        if df.empty:
            return pd.DataFrame(columns=["fecha", "notas"])

        df = df[["fecha", "notas"]].copy()
        df = df.sort_values("fecha", ascending=False)

        return df

class RegistroService:
    def __init__(self, repo: Repo, cfg: AppConfig):
        self.repo = repo
        self.cfg = cfg

    def suggest_doctor_names(self, doctor_name: Optional[str], n: int = 3, cutoff: float = 0.6) -> List[str]:
        """
        Devuelve una lista de nombres del config.json parecidos a `doctor_name`.
        Normaliza acentos/mayúsculas y usa difflib.get_close_matches.
        Si doctor_name es vacío o ya coincide perfectamente, retorna [].
        """
        if not doctor_name or str(doctor_name).strip() == "":
            return []
        name_norm = normalize_name(doctor_name)
        keys = list((self.cfg.doctors_percent or {}).keys())
        # mapa normalizado -> original
        norm_map = { normalize_name(k): k for k in keys if k is not None }
        norm_keys = list(norm_map.keys())
        if name_norm in norm_keys:
            return []
        matches = get_close_matches(name_norm, norm_keys, n=n, cutoff=cutoff)
        return [norm_map[m] for m in matches]

    def registrar_consulta(self, paciente: str, telefono: Optional[str], items_input: List[Dict], doctor: Optional[str], promotor: Optional[str], tecnico_por_item_default: Optional[str]) -> Tuple[str, List[str]]:
        # upsert paciente
        pid, tel_final, existed = self.repo.upsert_patient(paciente, telefono)
        # ensure personal/tratamientos and construct items
        items_prepared = []
        for it in items_input:
            # either id_tratamiento provided or get one by name
            if it.get("id_tratamiento"):
                tid = it["id_tratamiento"]
            else:
                tid = self.repo.ensure_tratamiento(it.get("descripcion_item",""))
            tup = {
                "id_tratamiento": tid,
                "descripcion_item": it.get("descripcion_item",""),
                "cantidad": float(it.get("cantidad",1)),
                "precio_unitario": float(it.get("precio_unitario",0)),
                "id_tecnico_item": it.get("id_tecnico_item") or (tecnico_por_item_default or ""),
                "observaciones_item": it.get("observaciones_item","")
            }
            items_prepared.append(tup)
        did = self.repo.ensure_personal(doctor, "doctor") if doctor else ""
        pmid = self.repo.ensure_personal(promotor, "promotor") if promotor else ""
        cid, created_item_ids = self.repo.add_consulta_with_items(id_paciente=pid, fecha=datetime.now(), id_doctor=did, id_promotor=pmid, items=items_prepared)
        return cid, created_item_ids

    def eliminar_personal(self, id_personal):
        """
        Elimina (lógicamente) el personal.
        Marca activo = 0 y registra en historial.
        """
        df = self.repo.dfs.get("Personal", pd.DataFrame())

        if id_personal not in df["id_personal"].values:
            raise ValueError(f"El personal {id_personal} no existe")

        # Marcar como inactivo
        df.loc[df["id_personal"] == id_personal, "activo"] = 0
        df.loc[df["id_personal"] == id_personal, "modificado_en"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Guardar
        self.repo.save()

        # Registrar historial
        self._registrar_historial_personal(id_personal, "eliminado")

    def _registrar_historial_personal(self, id_personal, accion):
        dfh = self.repo.df_hist_personal

        nuevo = {
            "id_personal": id_personal,
            "accion": accion,
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        dfh.loc[len(dfh)] = nuevo
        self.repo.save_historial_personal()

class ReporteService:
    def __init__(self, repo: Repo, cfg: AppConfig):
        self.repo = repo
        self.cfg = cfg

    def filtrar(self, desde: Optional[date], hasta: Optional[date], paciente: Optional[str], doctor: Optional[str], promotor: Optional[str], tecnico: Optional[str], tratamiento: Optional[str]) -> pd.DataFrame:
        df = self.repo.get_consultas_items_join()

        print("\n==== DEBUG 3: DF recibido por ReporteService.filtrar ====")
        print(df[["id_item","descripcion_item","tratamiento"]].head(20))

        # fechas
        if isinstance(desde, date):
            df = df[df["fecha"] >= pd.Timestamp(desde)]
        if isinstance(hasta, date):
            df = df[df["fecha"] < pd.Timestamp(hasta) + pd.Timedelta(days=1)]
        # filtros textuales (contiene)
        def contains(col, val):
            return df[col].fillna("").str.lower().str.contains((val or "").strip().lower(), na=False)
        if paciente: df = df[contains("paciente", paciente)]
        if doctor: df = df[contains("doc_nombre", doctor)]
        if promotor: df = df[contains("pro_nombre", promotor)]
        if tecnico: df = df[contains("tec_nombre", tecnico)]

        # --- FILTRO POR TRATAMIENTO ---
        if tratamiento:
            t = tratamiento.strip().lower()
            df = df[
                df["descripcion_item"].fillna("").str.lower().str.contains(t, na=False)
            ]
        # utilidad por item: aquí NO restamos comisiones - comisiones se calculan luego en corte
        df["utilidad_bruta_item"] = df["subtotal"] - 0.0
        df["tratamiento"] = df["id_tratamiento"].fillna("") + " - " + df["tratamiento"].fillna("")


        print("\n==== DEBUG 3B: Después de FORZAR tratamiento en filtrar ====")
        print(df[["id_item","descripcion_item","tratamiento"]].head(20))
        return df

class CorteService:
    def __init__(self, repo: Repo, cfg: AppConfig):
        self.repo = repo
        self.cfg = cfg

    def _recalcular_comisiones_para_corte(self, id_corte: str):
        """
        Recalcula comisiones para todas las filas de Corte_Detalle del id_corte.
        Hace merge con Consulta_Items para obtener subtotal y usa la lógica robusta
        de mapeo de doctores (por id_personal -> nombre, normalización y fallback).
        """
        # preparar mapas normalizados
        doctors_map_norm = {}
        try:
            for k, v in (self.cfg.doctors_percent or {}).items():
                if k is None:
                    continue
                doctors_map_norm[normalize_name(k)] = float(v)
        except Exception:
            doctors_map_norm = {}

        personal_df = self.repo.dfs.get("Personal", pd.DataFrame()).copy()
        id_to_name_norm = {}
        if not personal_df.empty:
            for _, r in personal_df.iterrows():
                pid = str(r.get("id_personal", "")).strip()
                if pid:
                    id_to_name_norm[pid] = normalize_name(r.get("nombre", ""))

        def pct_for_row(row):
            # 1) intentar por id_doctor si existe y mapea a nombre
            id_doc = str(row.get("id_doctor", "") or "").strip()
            if id_doc and id_doc in id_to_name_norm:
                return doctors_map_norm.get(id_to_name_norm[id_doc], 0.0)
            # 2) por nombre explícito
            n = normalize_name(row.get("doc_nombre", ""))
            if n and n in doctors_map_norm:
                return doctors_map_norm[n]
            # 3) fallback por coincidencia parcial
            for k in doctors_map_norm.keys():
                if (n and n.startswith(k)) or (k and k.startswith(n)) or (k in n) or (n in k):
                    return doctors_map_norm[k]
            return 0.0

        # cargar detalle del corte y items
        det_all = self.repo.dfs.get("Corte_Detalle", pd.DataFrame()).copy()
        rows = det_all[det_all["id_corte"].astype(str) == str(id_corte)].copy()
        if rows.empty:
            return

        items = self.repo.dfs.get("Consulta_Items", pd.DataFrame()).copy()
        # normalizar ids como strings
        for df in (rows, items):
            for c in ["id_item", "id_consulta"]:
                if c in df.columns:
                    df[c] = df[c].astype("string").fillna("")

        merged = rows.merge(items, on="id_item", how="left", suffixes=("_det", "_item"))

        # subtotal robusto
        merged["subtotal"] = pd.to_numeric(merged.get("subtotal", merged.get("subtotal_item", merged.get("subtotal_det", 0))), errors="coerce").fillna(0.0)

        _sub_vals = []
        for c in ("subtotal", "subtotal_item", "subtotal_det"):
            if c in merged.columns:
                _sub_vals.append(pd.to_numeric(merged[c], errors="coerce"))
        if _sub_vals:
            _sub_concat = pd.concat(_sub_vals, axis=1)
            merged["subtotal"] = _sub_concat.bfill(axis=1).iloc[:, 0].fillna(0.0).astype(float)
        else:
            merged["subtotal"] = 0.0

        # 2) asegurar columnas textuales que usaremos (evita AttributeError sobre .fillna())
        _text_cols = ["doc_nombre", "doctor", "id_doctor", "promotor", "pro_nombre", "tec_nombre"]
        for col in _text_cols:
            if col in merged.columns:
                try:
                    merged[col] = merged[col].astype("string").fillna("")
                except Exception:
                    # fallback seguro: convertir valor a string o cadena vacía si es NaN
                    merged[col] = merged[col].apply(lambda v: "" if pd.isna(v) else str(v))
            else:
                merged[col] = ""

        # Normalizar nombres: si no existe doc_nombre usar doctor (ya asegurado como string)
        if merged["doc_nombre"].eq("").all() and "doctor" in merged.columns:
            merged["doc_nombre"] = merged["doctor"]


        # calcular %
        merged["pct_doc"] = merged.apply(pct_for_row, axis=1)
        merged["comision_doctor_new"] = (merged["subtotal"].fillna(0.0) * merged["pct_doc"].fillna(0.0)).round(2)
        merged["comision_promotor_new"] = merged.apply(
            lambda r: round(float(r.get("subtotal", 0.0)) * float(self.cfg.promoter_percent or 0.0), 2)
            if str(r.get("promotor", "") or r.get("pro_nombre", "")).strip() else 0.0,
            axis=1
        )

        # aplicar valores en Corte_Detalle
        for _, r in merged.iterrows():
            id_det = r.get("id_corte_detalle")
            if not id_det:
                continue
            idx = self.repo.dfs["Corte_Detalle"].index[self.repo.dfs["Corte_Detalle"]["id_corte_detalle"] == id_det].tolist()
            if not idx:
                continue
            i = idx[0]
            self.repo.dfs["Corte_Detalle"].at[i, "comision_doctor"] = float(r.get("comision_doctor_new", 0.0) or 0.0)
            self.repo.dfs["Corte_Detalle"].at[i, "comision_promotor"] = float(r.get("comision_promotor_new", 0.0) or 0.0)

        # recalcular totales y guardar
        tot_ing = float(merged["subtotal"].sum())
        tot_cd = float(merged["comision_doctor_new"].sum())
        tot_cp = float(merged["comision_promotor_new"].sum())
        tot_ct = 0.0
        tot_util = round(tot_ing - tot_ct - tot_cd - tot_cp, 2)

        idxc = self.repo.dfs["Cortes"].index[self.repo.dfs["Cortes"]["id_corte"] == id_corte].tolist()
        if idxc:
            i = idxc[0]
            self.repo.dfs["Cortes"].at[i, "total_ingresos"] = tot_ing
            self.repo.dfs["Cortes"].at[i, "total_comisiones_doctores"] = tot_cd
            self.repo.dfs["Cortes"].at[i, "total_comisiones_promotores"] = tot_cp
            self.repo.dfs["Cortes"].at[i, "total_costo_tecnico"] = tot_ct
            self.repo.dfs["Cortes"].at[i, "total_utilidad"] = tot_util

        self.repo.save()

    def generar_corte(self, fecha: date) -> Tuple[str, pd.DataFrame, Dict[str, float]]:
        """
        Genera o actualiza un corte para la fecha dada.
        - Si NO existe corte para la fecha: crea uno nuevo con todos los items del día.
        - Si YA existe corte para la fecha: añade solo los items que no estén ya en Corte_Detalle para ese corte,
          y actualiza los totales del corte.
        Devuelve (id_corte, df_items_nuevos, resumen_actualizado).
        """
        rep = ReporteService(self.repo, self.cfg)
        # Trae todos los items del día (rango [fecha 00:00:00, fecha 23:59:59])
        df_items = rep.filtrar(fecha, fecha, None, None, None, None, None)
        if df_items.empty:
            # nada que cortar
            return "", pd.DataFrame(), {"total_ingresos": 0.0, "total_costo_tecnico": 0.0, "total_comisiones_doctores": 0.0, "total_comisiones_promotores": 0.0, "total_utilidad": 0.0}

        # Normalizar nombres de doctores para mapear comisiones (desde config)
        doctors_map_norm = {}
        try:
            for k, v in (self.cfg.doctors_percent or {}).items():
                if k is None:
                    continue
                doctors_map_norm[normalize_name(k)] = float(v)
        except Exception:
            doctors_map_norm = {}

        def doctor_pct_norm(name):
            n = normalize_name(name)
            return float(doctors_map_norm.get(n, 0.0))

        # Calcula comisiones por item (aplicando mapping normalizado)
        df_items = df_items.copy()
        df_items["pct_doc"] = df_items["doc_nombre"].fillna("").apply(lambda n: doctor_pct_norm(n))

        df_items["comision_doctor"] = (df_items["subtotal"].fillna(0.0) * df_items["pct_doc"]).round(2)
        df_items["comision_promotor"] = df_items.apply(
            lambda r: round(float(r["subtotal"] or 0) * float(self.cfg.promoter_percent or 0.0), 2)
            if pd.notna(r.get("pro_nombre")) and str(r.get("pro_nombre")).strip() != "" else 0.0, axis=1
        )

        # Totales calculados sobre los items (por ahora costo_tecnico por item no está en schema -> 0)
        tot_ing = float(df_items["subtotal"].sum())
        tot_cd = float(df_items["comision_doctor"].sum())
        tot_cp = float(df_items["comision_promotor"].sum())
        tot_ct = 0.0
        tot_util = round(tot_ing - tot_ct - tot_cd - tot_cp, 2)

        # Checa si ya existe un corte para la fecha
        cortes_df = self.repo.dfs["Cortes"]
        fecha_iso = pd.Timestamp(fecha).normalize()
        # Buscamos coincidencia exacta en fecha_corte (date or datetime allowed)
        mask_fecha = cortes_df["fecha_corte"].apply(lambda x: pd.to_datetime(x).normalize() if pd.notna(x) and str(x).strip() != "" else pd.NaT)
        existentes = cortes_df[mask_fecha == fecha_iso]

        # Items que ya están en Corte_Detalle (para evitar duplicados)
        detalle_df_all = self.repo.dfs.get("Corte_Detalle", pd.DataFrame())
        existing_items_in_any_corte = set(detalle_df_all["id_item"].fillna("").astype(str).tolist()) if not detalle_df_all.empty else set()

        if not existentes.empty:
            # Tomamos el corte existente (si hay varios, tomamos el primero; idealmente sólo habrá uno por día)
            row = existentes.iloc[0]
            id_corte = row["id_corte"]
            # Filtramos los items del día que NO estén ya registrados en Corte_Detalle para ese corte
            detalle_actual = self.repo.dfs["Corte_Detalle"]
            items_ya_en_este_corte = set(detalle_actual[detalle_actual["id_corte"].astype(str) == str(id_corte)]["id_item"].fillna("").astype(str).tolist())

            # A veces migración no llenó id_item: en ese caso la comparación por id_item puede no filtrar; asumimos id_item preferente
            df_items_nuevos = df_items[~df_items["id_item"].astype(str).isin(items_ya_en_este_corte)].copy()

            if df_items_nuevos.empty:
                # No hay items nuevos desde el último corte: pero recalculamos por si hay filas previas con comisiones incorrectas
                try:
                    self._recalcular_comisiones_para_corte(id_corte)
                except Exception as e:
                    # no queremos que una excepción bloquee la devolución; la registramos en consola/log
                    print("Warning: error al recalcular comisiones para corte existente:", e)

                # leer la fila actualizada del corte y devolver resumen
                row = self.repo.dfs["Cortes"].loc[self.repo.dfs["Cortes"]["id_corte"] == id_corte].iloc[0]
                resumen = {
                    "total_ingresos": float(row.get("total_ingresos", 0.0)),
                    "total_costo_tecnico": float(row.get("total_costo_tecnico", 0.0)),
                    "total_comisiones_doctores": float(row.get("total_comisiones_doctores", 0.0)),
                    "total_comisiones_promotores": float(row.get("total_comisiones_promotores", 0.0)),
                    "total_utilidad": float(row.get("total_utilidad", 0.0)),
                }
                return id_corte, pd.DataFrame(), resumen

            # calculos sólo sobre los items nuevos
            add_ing = float(df_items_nuevos["subtotal"].sum())
            add_cd = float(df_items_nuevos["comision_doctor"].sum())
            add_cp = float(df_items_nuevos["comision_promotor"].sum())
            add_ct = 0.0
            add_util = round(add_ing - add_ct - add_cd - add_cp, 2)

            # Agregar filas en Corte_Detalle para los items nuevos
            det_df = self.repo.dfs["Corte_Detalle"]
            filas = []
            for _, r in df_items_nuevos.iterrows():
                filas.append({
                    "id_corte_detalle": _new_id("L", det_df["id_corte_detalle"]) if not det_df.empty else _new_id("L", pd.Series([])),
                    "id_corte": id_corte,
                    "id_consulta": r.get("id_consulta", ""),
                    "id_item": r.get("id_item", "") if "id_item" in r.index else "",
                    "doctor": r.get("doc_nombre", ""),
                    "promotor": r.get("pro_nombre", ""),
                    "tecnico": r.get("tec_nombre", ""),
                    "comision_doctor": float(r.get("comision_doctor", 0.0)),
                    "comision_promotor": float(r.get("comision_promotor", 0.0)),
                    "costo_tecnico": 0.0,
                    "incluido": 1
                })
            self.repo.dfs["Corte_Detalle"] = pd.concat([det_df, pd.DataFrame(filas)], ignore_index=True)

            # Actualizar totales en la fila del corte
            idx = self.repo.dfs["Cortes"].index[self.repo.dfs["Cortes"]["id_corte"] == id_corte].tolist()
            if idx:
                i = idx[0]
                # sumar los añadidos
                self.repo.dfs["Cortes"].at[i, "total_ingresos"] = float(self.repo.dfs["Cortes"].at[i, "total_ingresos"] or 0.0) + add_ing
                self.repo.dfs["Cortes"].at[i, "total_comisiones_doctores"] = float(self.repo.dfs["Cortes"].at[i, "total_comisiones_doctores"] or 0.0) + add_cd
                self.repo.dfs["Cortes"].at[i, "total_comisiones_promotores"] = float(self.repo.dfs["Cortes"].at[i, "total_comisiones_promotores"] or 0.0) + add_cp
                self.repo.dfs["Cortes"].at[i, "total_costo_tecnico"] = float(self.repo.dfs["Cortes"].at[i, "total_costo_tecnico"] or 0.0) + add_ct
                # recalcula utilidad total
                new_tot_ing = float(self.repo.dfs["Cortes"].at[i, "total_ingresos"])
                new_tot_cd = float(self.repo.dfs["Cortes"].at[i, "total_comisiones_doctores"])
                new_tot_cp = float(self.repo.dfs["Cortes"].at[i, "total_comisiones_promotores"])
                new_tot_ct = float(self.repo.dfs["Cortes"].at[i, "total_costo_tecnico"])
                self.repo.dfs["Cortes"].at[i, "total_utilidad"] = round(new_tot_ing - new_tot_ct - new_tot_cd - new_tot_cp, 2)
            # persistir
            self.repo.save()
            # Armar resumen actualizado para devolver
            row_upd = self.repo.dfs["Cortes"].loc[self.repo.dfs["Cortes"]["id_corte"] == id_corte].iloc[0]
            resumen = {
                "total_ingresos": float(row_upd.get("total_ingresos", 0.0)),
                "total_costo_tecnico": float(row_upd.get("total_costo_tecnico", 0.0)),
                "total_comisiones_doctores": float(row_upd.get("total_comisiones_doctores", 0.0)),
                "total_comisiones_promotores": float(row_upd.get("total_comisiones_promotores", 0.0)),
                "total_utilidad": float(row_upd.get("total_utilidad", 0.0)),
            }
            return id_corte, df_items_nuevos, resumen

        else:
            # No existe corte para la fecha: creamos uno nuevo con todos los items del día
            cortes = self.repo.dfs["Cortes"]
            id_corte = _new_id("D", cortes["id_corte"])
            now = datetime.now().isoformat(timespec="seconds")
            resumen_row = {
                "id_corte": id_corte,
                "fecha_corte": fecha.isoformat(),
                "rango_inicio": f"{fecha} 00:00:00",
                "rango_fin": f"{fecha} 23:59:59",
                "total_ingresos": tot_ing,
                "total_costo_tecnico": tot_ct,
                "total_comisiones_doctores": tot_cd,
                "total_comisiones_promotores": tot_cp,
                "total_utilidad": tot_util,
                "creado_por": "sistema",
                "creado_en": now,
                "version": 1
            }
            self.repo.dfs["Cortes"] = pd.concat([cortes, pd.DataFrame([resumen_row])], ignore_index=True)

            # Detalle por item -> Corte_Detalle
            det = self.repo.dfs["Corte_Detalle"]
            filas = []
            for _, r in df_items.iterrows():
                filas.append({
                    "id_corte_detalle": _new_id("L", det["id_corte_detalle"]) if not det.empty else _new_id("L", pd.Series([])),
                    "id_corte": id_corte,
                    "id_consulta": r.get("id_consulta", ""),
                    "id_item": r.get("id_item", "") if "id_item" in r.index else "",
                    "doctor": r.get("doc_nombre", ""),
                    "promotor": r.get("pro_nombre", ""),
                    "tecnico": r.get("tec_nombre", ""),
                    "comision_doctor": float(r.get("comision_doctor", 0.0)),
                    "comision_promotor": float(r.get("comision_promotor", 0.0)),
                    "costo_tecnico": 0.0,
                    "incluido": 1
                })
            self.repo.dfs["Corte_Detalle"] = pd.concat([det, pd.DataFrame(filas)], ignore_index=True)
            self.repo.save()
            # recalculamos para asegurar comisiones correctas (por si hubo problemas de normalización)
            self._recalcular_comisiones_para_corte(id_corte)
            resumen = {"total_ingresos": tot_ing, "total_costo_tecnico": tot_ct, "total_comisiones_doctores": tot_cd, "total_comisiones_promotores": tot_cp, "total_utilidad": tot_util}
            return id_corte, df_items, resumen

    def detalle_corte(self, id_corte: str):
        """
        Devuelve (detalle_df, resumen_dict) para un corte.
        Esta versión ya está corregida para mostrar siempre el tratamiento
        usando descripcion_item como fuente principal.
        """
        import numpy as np

        det = self.repo.dfs.get("Corte_Detalle", pd.DataFrame()).copy()
        if det.empty:
            return pd.DataFrame(), {"doctores": {}, "promotores": {}, "tecnicos": {}}

        # Filtrar por corte
        det["id_corte"] = det["id_corte"].astype("string").fillna("")
        det = det[det["id_corte"] == str(id_corte)]
        if det.empty:
            return pd.DataFrame(), {"doctores": {}, "promotores": {}, "tecnicos": {}}

        # Tablas auxiliares
        items = self.repo.dfs.get("Consulta_Items", pd.DataFrame()).copy()
        cons = self.repo.dfs.get("Consultas", pd.DataFrame()).copy()
        pac = self.repo.dfs.get("Pacientes", pd.DataFrame())[["id_paciente","nombre"]] \
                .rename(columns={"nombre":"paciente"}) if "Pacientes" in self.repo.dfs else pd.DataFrame(columns=["id_paciente","paciente"])
        trat = self.repo.dfs.get("Tratamientos", pd.DataFrame())[["id_tratamiento","nombre"]] \
                .rename(columns={"nombre":"tratamiento"}) if "Tratamientos" in self.repo.dfs else pd.DataFrame(columns=["id_tratamiento","tratamiento"])

        # Normalizar tipos
        for df, cols in [
            (det, ["id_item","id_consulta"]),
            (items, ["id_item","id_consulta"]),
            (cons, ["id_consulta","id_paciente"]),
            (pac, ["id_paciente"]),
            (trat, ["id_tratamiento"])
        ]:
            for c in cols:
                if c in df.columns:
                    df[c] = df[c].astype("string").fillna("")

        # Merge principal (preferimos unir por id_item)
        if "id_item" in det.columns and not det["id_item"].astype(str).replace("", pd.NA).isna().all():
            df = det.merge(items, on="id_item", how="left", suffixes=("_det", "_item"))
        else:
            # fallback: unir por id_consulta
            if "id_consulta" in det.columns and det["id_consulta"].astype(str).str.strip().any():
                df = det.merge(cons, on="id_consulta", how="left", suffixes=("_det","_cons"))
            else:
                df = det.copy()

        # Si falta id_consulta, intentar rescatarlo desde items
        if "id_consulta" not in df.columns and "id_consulta" in items.columns:
            df = df.merge(items[["id_item","id_consulta"]], on="id_item", how="left")

        # Unir paciente
        if "id_consulta" in df.columns and "id_paciente" not in df.columns:
            if not cons.empty:
                df = df.merge(cons[["id_consulta","id_paciente"]], on="id_consulta", how="left")
        if "id_paciente" in df.columns:
            df = df.merge(pac, on="id_paciente", how="left")

        # ---------------------------
        # 🔧 *** ARREGLO CLAVE ***
        # Unificar descripcion_item aunque venga con sufijos
        # ---------------------------
        if "descripcion_item" not in df.columns:
            for col in ["descripcion_item_item", "descripcion_item_det"]:
                if col in df.columns:
                    df["descripcion_item"] = df[col]
                    break
            else:
                df["descripcion_item"] = ""

        # ---------------------------
        # Unificar id_tratamiento si existe
        # ---------------------------
        if "id_tratamiento" in df.columns:
            df = df.merge(trat, on="id_tratamiento", how="left")

        # Unificar numéricos
        def coalesce_numeric(frame, targets, out):
            vals = []
            for t in targets:
                if t in frame.columns:
                    vals.append(pd.to_numeric(frame[t], errors="coerce"))
            if vals:
                frame[out] = pd.concat(vals, axis=1).bfill(axis=1).iloc[:,0].fillna(0.0)
            else:
                frame[out] = 0.0
            return frame

        df = coalesce_numeric(df, ["subtotal","subtotal_item","subtotal_det"], "subtotal")
        df = coalesce_numeric(df, ["precio_unitario","precio_unitario_item"], "precio_unitario")
        df = coalesce_numeric(df, ["cantidad","cantidad_item"], "cantidad")

        # Comisiones
        for c in ("comision_doctor","comision_promotor","costo_tecnico"):
            df[c] = pd.to_numeric(df.get(c, 0.0), errors="coerce").fillna(0.0)

        if "fecha" in df.columns:
            df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")

        # Normalizar personal
        if "doctor" in df.columns and "doc_nombre" not in df.columns:
            df["doc_nombre"] = df["doctor"]
        if "promotor" in df.columns and "pro_nombre" not in df.columns:
            df["pro_nombre"] = df["promotor"]
        if "tecnico" in df.columns and "tec_nombre" not in df.columns:
            df["tec_nombre"] = df["tecnico"]

        # ---------------------------
        # 🔥 EL PASO CRÍTICO:
        # tratamiento SIEMPRE es descripcion_item
        # ---------------------------
        df["tratamiento"] = df["tratamiento"].fillna(df["descripcion_item"])

        df["utilidad_fila"] = (
            df["subtotal"] - df["costo_tecnico"] - df["comision_doctor"] - df["comision_promotor"]
        ).round(2)

        # Resumen por persona
        por_doc = df.groupby("doc_nombre")["comision_doctor"].sum().round(2).to_dict()
        por_pro = df.groupby("pro_nombre")["comision_promotor"].sum().round(2).to_dict()
        por_tec = df.groupby("tec_nombre")["costo_tecnico"].sum().round(2).to_dict()

        resumen = {"doctores": por_doc, "promotores": por_pro, "tecnicos": por_tec}

        # Columnas esperadas
        out_cols = [
            "fecha","paciente","tratamiento","descripcion_item","doc_nombre",
            "pro_nombre","tec_nombre","subtotal","precio_unitario","cantidad",
            "comision_doctor","comision_promotor","costo_tecnico","utilidad_fila"
        ]
        for c in out_cols:
            if c not in df.columns:
                df[c] = ""

        df_out = df[out_cols].copy()

        return df_out, resumen

class ComisionService:
    """
    Servicio de apoyo para la pestaña *Comisiones* de la GUI.
    NO modifica nada del flujo de consultas/cortes; sólo lee Consultas/Pacientes/Personal
    y escribe en la hoja Comisiones_Abonos.
    """
    def __init__(self, repo: Repo, cfg: AppConfig):
        self.repo = repo
        self.cfg = cfg

    # ------------------------
    # Helpers internos
    # ------------------------
    def _build_maps_personal(self):
        personal = self.repo.dfs.get("Personal", pd.DataFrame()).copy()
        if personal.empty:
            personal = pd.DataFrame(columns=["id_personal","nombre","rol","pct_comision","activo"])
        if "pct_comision" not in personal.columns:
            personal["pct_comision"] = 0.0
        personal["id_personal"] = personal.get("id_personal","").astype(str)
        personal["rol"] = personal.get("rol","").astype(str)

        pct_por_id = {}
        for _, r in personal.iterrows():
            pid = str(r.get("id_personal","")).strip()
            try:
                pct = float(r.get("pct_comision") or 0.0)
            except Exception:
                pct = 0.0
            pct_por_id[pid] = pct

        # mapa de doctores desde config.json
        cfg_doctors = {}
        try:
            for k, v in (self.cfg.doctors_percent or {}).items():
                if k is None:
                    continue
                cfg_doctors[normalize_name(k)] = float(v)
        except Exception:
            cfg_doctors = {}

        return personal, pct_por_id, cfg_doctors

    def _pct_doctor(self, id_doctor: str, nombre_doctor: str, pct_por_id: dict, cfg_doctors: dict) -> float:
        """
        Regla de negocio para % de doctor:
        1) Si en Personal.pct_comision hay valor > 0 -> usarlo.
        2) Si no, buscar en config.doctors_percent por nombre normalizado.
        3) Si nada, 0.0
        """
        pid = str(id_doctor or "").strip()
        if pid and pid in pct_por_id and pct_por_id[pid] > 0:
            return float(pct_por_id[pid])
        n = normalize_name(nombre_doctor or "")
        if n in cfg_doctors:
            return float(cfg_doctors[n])
        # fallback por coincidencia parcial
        for k in cfg_doctors.keys():
            if n and (n.startswith(k) or k.startswith(n) or k in n or n in k):
                return float(cfg_doctors[k])
        return 0.0

    # ---------------------------------------------------
    # 1) Resumen de comisiones de un doctor en un día
    # ---------------------------------------------------
    def resumen_dia_por_doctor(self, fecha: date, nombre_doctor: str):
        """
        Devuelve (df_resumen, totales_dict) para la tabla de:
        Hora | Paciente | Total | Metodo | Moneda | Comision | Pagado | Pendiente
        """
        cons = self.repo.dfs.get("Consultas", pd.DataFrame()).copy()
        if cons.empty:
            cols = ["hora","paciente","total","metodo","moneda","comision","pagado","pendiente"]
            return pd.DataFrame(columns=cols), {"total_generado":0.0, "pagado":0.0, "pendiente":0.0}

        cons["fecha"] = pd.to_datetime(cons["fecha"], errors="coerce")
        cons_dia = cons[cons["fecha"].dt.date == fecha].copy()
        if cons_dia.empty:
            cols = ["hora","paciente","total","metodo","moneda","comision","pagado","pendiente"]
            return pd.DataFrame(columns=cols), {"total_generado":0.0, "pagado":0.0, "pendiente":0.0}

        personal, pct_por_id, cfg_doctors = self._build_maps_personal()
        doctores = personal[personal["rol"].str.lower() == "doctor"].copy()
        doctores["nombre_norm"] = doctores.get("nombre","").fillna("").apply(normalize_name)
        target_norm = normalize_name(nombre_doctor)

        # IDs de este doctor
        ids_doctor = doctores.loc[doctores["nombre_norm"] == target_norm, "id_personal"].astype(str).tolist()
        if not ids_doctor:
            # fallback contains
            ids_doctor = doctores.loc[
                doctores.get("nombre","").fillna("").str.contains(nombre_doctor, case=False, na=False),
                "id_personal"
            ].astype(str).tolist()

        cons_dia["id_doctor"] = cons_dia.get("id_doctor","").astype(str)
        if ids_doctor:
            cons_dia = cons_dia[cons_dia["id_doctor"].isin(ids_doctor)]
        if cons_dia.empty:
            cols = ["hora","paciente","total","metodo","moneda","comision","pagado","pendiente"]
            return pd.DataFrame(columns=cols), {"total_generado":0.0, "pagado":0.0, "pendiente":0.0}

        # Join con Pacientes para nombre
        pac = self.repo.dfs.get("Pacientes", pd.DataFrame()).copy()
        if not pac.empty and "id_paciente" in pac.columns:
            pac = pac[["id_paciente","nombre"]].rename(columns={"nombre":"paciente"})
            cons_dia = cons_dia.merge(pac, on="id_paciente", how="left")

        # Campos base
        cons_dia["hora"] = cons_dia["fecha"].dt.strftime("%H:%M").fillna("")
        cons_dia["total_consulta"] = pd.to_numeric(cons_dia.get("total_consulta", 0.0), errors="coerce").fillna(0.0)
        cons_dia["metodo"] = cons_dia.get("metodo_pago","").fillna("")
        cons_dia["moneda"] = cons_dia.get("moneda","").fillna("")

        # Nombre del doctor para cada fila
        cons_dia = cons_dia.merge(
            doctores[["id_personal","nombre"]].rename(columns={"id_personal":"id_doc_personal","nombre":"nombre_doctor"}),
            left_on="id_doctor", right_on="id_doc_personal", how="left"
        )
        cons_dia["nombre_doctor"] = cons_dia["nombre_doctor"].fillna(nombre_doctor)

        # % comisión por fila
        cons_dia["pct_comision"] = cons_dia.apply(
            lambda r: self._pct_doctor(
                r.get("id_doctor",""),
                r.get("nombre_doctor",""),
                pct_por_id,
                cfg_doctors
            ),
            axis=1
        )
        cons_dia["comision"] = (cons_dia["total_consulta"] * cons_dia["pct_comision"]).round(2)

        # Abonos acumulados para (doctor, consulta)
        ab = self.repo.dfs.get("Comisiones_Abonos", pd.DataFrame()).copy()
        if not ab.empty:
            ab["id_personal"] = ab.get("id_personal","").astype(str)
            ab["id_consulta"] = ab.get("id_consulta","").astype(str)
            grp = ab.groupby(["id_personal","id_consulta"])["monto"].sum().reset_index().rename(columns={"monto":"pagado"})
            grp["pagado"] = pd.to_numeric(grp["pagado"], errors="coerce").fillna(0.0)
        else:
            grp = pd.DataFrame(columns=["id_personal","id_consulta","pagado"])

        cons_dia["id_personal"] = cons_dia["id_doctor"].astype(str)
        cons_dia["id_consulta"] = cons_dia["id_consulta"].astype(str)
        cons_dia = cons_dia.merge(grp, on=["id_personal","id_consulta"], how="left")
        cons_dia["pagado"] = pd.to_numeric(cons_dia.get("pagado",0.0), errors="coerce").fillna(0.0)
        cons_dia["pendiente"] = (cons_dia["comision"] - cons_dia["pagado"]).round(2)
        cons_dia.loc[cons_dia["pendiente"] < 0, "pendiente"] = 0.0

        cols = ["hora","paciente","total_consulta","metodo","moneda","comision","pagado","pendiente"]
        for c in cols:
            if c not in cons_dia.columns:
                cons_dia[c] = "" if c in ("hora","paciente","metodo","moneda") else 0.0

        out = cons_dia[cols].rename(columns={"total_consulta":"total"})

        total_generado = float(out["comision"].sum())
        pagado = float(out["pagado"].sum())
        pendiente = float(out["pendiente"].sum())
        totales = {
            "total_generado": round(total_generado,2),
            "pagado": round(pagado,2),
            "pendiente": round(pendiente,2)
        }
        return out, totales

    # ---------------------------------------------------
    # 2) Pendientes por personal (doctor/promotor)
    # ---------------------------------------------------
    def pendientes_por_personal(self, id_personal: str) -> pd.DataFrame:
        """
        Devuelve DataFrame para la tabla de 'Abonos a Comisiones' de la GUI:
        Fecha | Paciente | Total | Comision | Abonado | Restante | Metodo | Moneda
        Sólo filas donde Restante > 0.
        """
        cons = self.repo.dfs.get("Consultas", pd.DataFrame()).copy()
        if cons.empty:
            cols = ["fecha","paciente","total","comision","abonado","restante","metodo","moneda"]
            return pd.DataFrame(columns=cols)

        cons["fecha"] = pd.to_datetime(cons["fecha"], errors="coerce")
        cons["id_doctor"] = cons.get("id_doctor","").astype(str)
        cons["id_promotor"] = cons.get("id_promotor","").astype(str)
        idp = str(id_personal)

        cons = cons[(cons["id_doctor"] == idp) | (cons["id_promotor"] == idp)].copy()
        if cons.empty:
            cols = ["fecha","paciente","total","comision","abonado","restante","metodo","moneda"]
            return pd.DataFrame(columns=cols)

        # Join pacientes
        pac = self.repo.dfs.get("Pacientes", pd.DataFrame()).copy()
        if not pac.empty and "id_paciente" in pac.columns:
            pac = pac[["id_paciente","nombre"]].rename(columns={"nombre":"paciente"})
            cons = cons.merge(pac, on="id_paciente", how="left")

        cons["total_consulta"] = pd.to_numeric(cons.get("total_consulta",0.0), errors="coerce").fillna(0.0)
        cons["metodo"] = cons.get("metodo_pago","").fillna("")
        cons["moneda"] = cons.get("moneda","").fillna("")

        # maps para % doctor
        personal, pct_por_id, cfg_doctors = self._build_maps_personal()

        def pct_row(r):
            if str(r.get("id_doctor","")) == idp:
                # es doctor
                nombre_doc = ""
                if not personal.empty:
                    m = personal[personal["id_personal"].astype(str) == idp]
                    if not m.empty:
                        nombre_doc = m.iloc[0].get("nombre","")
                return self._pct_doctor(idp, nombre_doc, pct_por_id, cfg_doctors)
            if str(r.get("id_promotor","")) == idp:
                # es promotor
                try:
                    return float(self.cfg.promoter_percent or 0.0)
                except Exception:
                    return 0.0
            return 0.0

        cons["pct_comision"] = cons.apply(pct_row, axis=1)
        cons["comision"] = (cons["total_consulta"] * cons["pct_comision"]).round(2)

        # Abonos acumulados
        ab = self.repo.dfs.get("Comisiones_Abonos", pd.DataFrame()).copy()
        if not ab.empty:
            ab["id_personal"] = ab.get("id_personal","").astype(str)
            ab["id_consulta"] = ab.get("id_consulta","").astype(str)
            grp = ab.groupby(["id_personal","id_consulta"])["monto"].sum().reset_index().rename(columns={"monto":"abonado"})
            grp["abonado"] = pd.to_numeric(grp["abonado"], errors="coerce").fillna(0.0)
        else:
            grp = pd.DataFrame(columns=["id_personal","id_consulta","abonado"])

        cons["id_personal"] = idp
        cons["id_consulta"] = cons["id_consulta"].astype(str)
        cons = cons.merge(grp, on=["id_personal","id_consulta"], how="left")
        cons["abonado"] = pd.to_numeric(cons.get("abonado",0.0), errors="coerce").fillna(0.0)
        cons["restante"] = (cons["comision"] - cons["abonado"]).round(2)
        cons = cons[cons["restante"] > 0.0]

        cons["fecha"] = cons["fecha"].dt.date

        cols = ["fecha","paciente","total_consulta","comision","abonado","restante","metodo","moneda"]
        for c in cols:
            if c not in cons.columns:
                cons[c] = "" if c in ("paciente","metodo","moneda") else 0.0

        out = cons[cols].rename(columns={"total_consulta":"total"})
        out = out.sort_values("fecha")
        return out

    # ---------------------------------------------------
    # 3) Registrar abono (por IDs)
    # ---------------------------------------------------
    def registrar_abono_ids(self, id_personal: str, id_consulta: str, monto: float,
                            moneda: str = "", nota: str = "", fecha_pago: datetime | None = None) -> str:
        """
        Crea una fila en Comisiones_Abonos usando IDs directos.
        Devuelve id_abono.
        """
        if fecha_pago is None:
            fecha_pago = datetime.now()
        df = self.repo.dfs.get("Comisiones_Abonos", pd.DataFrame(columns=SHEETS["Comisiones_Abonos"])).copy()
        # generar id_abono con _new_id
        if "id_abono" not in df.columns:
            df["id_abono"] = ""
        aid = _new_id("A", df["id_abono"]) if not df.empty else _new_id("A", pd.Series([]))

        nueva = {
            "id_abono": aid,
            "id_personal": str(id_personal),
            "id_consulta": str(id_consulta),
            "fecha_pago": fecha_pago,
            "monto": float(monto),
            "moneda": moneda or "",
            "nota": nota or ""
        }
        df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
        self.repo.dfs["Comisiones_Abonos"] = df
        self.repo.save()
        return aid

    # ---------------------------------------------------
    # 4) Registrar abono pensando en la GUI actual
    # ---------------------------------------------------
    def registrar_abono_por_nombre(self, personal_nombre: str, fecha_consulta: date,
                                   paciente_nombre: str, monto: float, nota: str = "") -> str | None:
        """
        Versión helper para la GUI actual, que sólo manda:
        - nombre del personal,
        - fecha (de la consulta) y
        - nombre del paciente.
        Busca la consulta correspondiente y crea el abono.
        Devuelve id_abono o None si no encontró coincidencias.
        """
        personal = self.repo.dfs.get("Personal", pd.DataFrame()).copy()
        if personal.empty:
            return None
        personal["nombre_norm"] = personal.get("nombre","").fillna("").apply(normalize_name)
        target_norm = normalize_name(personal_nombre)
        print("[DEBUG] registrar_abono_por_nombre -> personal_target_norm:", target_norm)
        match = personal[personal["nombre_norm"] == target_norm]
        if match.empty:
            match = personal[personal.get("nombre","").fillna("").str.contains(personal_nombre, case=False, na=False)]
        if match.empty:
            return None
        per_row = match.iloc[0]
        id_personal = str(per_row["id_personal"])

        cons = self.repo.dfs.get("Consultas", pd.DataFrame()).copy()
        if cons.empty:
            return None
        cons["fecha"] = pd.to_datetime(cons["fecha"], errors="coerce")
        cons["fecha_dia"] = cons["fecha"].dt.date

        pac = self.repo.dfs.get("Pacientes", pd.DataFrame()).copy()
        if pac.empty:
            return None
        pac["nombre_norm"] = pac.get("nombre","").fillna("").apply(normalize_name)

        cons = cons.merge(pac[["id_paciente","nombre_norm"]], on="id_paciente", how="left")
        print("[DEBUG] Consultas tras merge pac:", cons[["id_consulta","fecha_dia","nombre_norm","id_doctor","id_promotor"]].head(10))

        paciente_norm = normalize_name(paciente_nombre)
        print("[DEBUG] Buscando por fecha/paciente -> fecha_consulta:", fecha_consulta, "paciente_norm:", paciente_norm,
              "filas totales:", len(cons))
        cons = cons[cons["fecha_dia"] == fecha_consulta]
        cons = cons[cons["nombre_norm"] == paciente_norm]
        print("[DEBUG] Filtro por fecha/paciente -> filas:", len(cons))
        cons["id_doctor"] = cons.get("id_doctor","").astype(str)
        cons["id_promotor"] = cons.get("id_promotor","").astype(str)

        cons = cons[(cons["id_doctor"] == id_personal) | (cons["id_promotor"] == id_personal)]
        print("[DEBUG] Filtro por personal -> filas:", len(cons), "id_personal:", id_personal)
        if cons.empty:
            return None

        # si hay varias, tomar la más reciente
        cons = cons.sort_values("fecha", ascending=False)
        row = cons.iloc[0]
        id_consulta = str(row["id_consulta"])
        moneda = str(row.get("moneda","") or "")

        return self.registrar_abono_ids(id_personal=id_personal, id_consulta=id_consulta,
                                        monto=float(monto), moneda=moneda, nota=nota)

    # ---------------------------------------------------
    # 5) Historial de abonos (para la tabla inferior)
    # ---------------------------------------------------
    def historial_abonos(self) -> pd.DataFrame:
        """
        Devuelve DataFrame con columnas:
        fecha_pago | personal | paciente | consulta | monto | moneda | nota
        ya unido con hojas Personal, Pacientes y Consultas.
        """
        ab = self.repo.dfs.get("Comisiones_Abonos", pd.DataFrame()).copy()
        if ab.empty:
            cols = ["fecha_pago","personal","paciente","consulta","monto","moneda","nota"]
            return pd.DataFrame(columns=cols)

        pers = self.repo.dfs.get("Personal", pd.DataFrame()).copy()
        cons = self.repo.dfs.get("Consultas", pd.DataFrame()).copy()
        pac = self.repo.dfs.get("Pacientes", pd.DataFrame()).copy()

        for df, cols_ids in [
            (ab, ["id_personal","id_consulta"]),
            (pers, ["id_personal"]),
            (cons, ["id_consulta","id_paciente"]),
            (pac, ["id_paciente"])
        ]:
            for c in cols_ids:
                if c in df.columns:
                    df[c] = df[c].astype(str)

        # join personal
        if not pers.empty and "id_personal" in pers.columns:
            ab = ab.merge(
                pers[["id_personal","nombre"]].rename(columns={"nombre":"personal"}),
                on="id_personal", how="left"
            )
        # join consulta -> paciente
        if not cons.empty and "id_consulta" in cons.columns:
            ab = ab.merge(cons[["id_consulta","id_paciente"]], on="id_consulta", how="left")
        if not pac.empty and "id_paciente" in pac.columns:
            ab = ab.merge(
                pac[["id_paciente","nombre"]].rename(columns={"nombre":"paciente"}),
                on="id_paciente", how="left"
            )

        ab["fecha_pago"] = pd.to_datetime(ab.get("fecha_pago"), errors="coerce")
        ab = ab.rename(columns={"id_consulta":"consulta"})
        ab["monto"] = pd.to_numeric(ab.get("monto",0.0), errors="coerce").fillna(0.0)
        ab["moneda"] = ab.get("moneda","").fillna("")
        ab["nota"] = ab.get("nota","").fillna("")

        cols = ["fecha_pago","personal","paciente","consulta","monto","moneda","nota"]
        for c in cols:
            if c not in ab.columns:
                ab[c] = "" if c in ("personal","paciente","consulta","moneda","nota") else 0.0

        ab = ab[cols].sort_values("fecha_pago", ascending=False)
        return ab
  

class KPIService:
    
    def __init__(self, repo: Repo, cfg: AppConfig):
        self.repo = repo
        self.cfg = cfg

    def calcular(self, desde: date, hasta: date) -> Dict:
        """
        Calcula KPIs entre dos fechas (inclusive).
        Robusto contra columnas faltantes, NaN y tipos raros.
        """
        rep = ReporteService(self.repo, self.cfg)
        df = rep.filtrar(desde, hasta, None, None, None, None, None).copy()

        # ==============================
        # Normalización de columnas base
        # ==============================

        # descripcion_item siempre presente y sin "nan" texto
        if "descripcion_item" not in df.columns:
            df["descripcion_item"] = ""
        df["descripcion_item"] = df["descripcion_item"].fillna("").astype(str)

        # --- TRATAMIENTO ---
        # 1) Usar la columna tratamiento que ya viene del Reporte (id + nombre)
        if "tratamiento" in df.columns:
            df["tratamiento"] = df["tratamiento"].fillna("").astype(str)
            # limpiar textos "nan", "None", etc.
            df["tratamiento"] = df["tratamiento"].replace(["nan", "None", "NONE", "NaN"], "")
        else:
            df["tratamiento"] = ""

        # 2) Si después de esto TODOS están vacíos, caer a descripcion_item
        if df["tratamiento"].str.strip().eq("").all():
            df["tratamiento"] = df["descripcion_item"].fillna("").astype(str)

        # --- Fechas ---
        if "fecha" in df.columns:
            df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")

        # --- Numéricos ---
        df["subtotal"] = pd.to_numeric(df.get("subtotal", 0.0), errors="coerce").fillna(0.0)
        df["cantidad"] = pd.to_numeric(df.get("cantidad", 1.0), errors="coerce").fillna(1.0)
        df["precio_unitario"] = pd.to_numeric(df.get("precio_unitario", 0.0), errors="coerce").fillna(0.0)

        # ==============================
        # Ingresos, pacientes, ticket
        # ==============================
        ingresos = float(df["subtotal"].sum())

        if "id_paciente" in df.columns:
            pacientes_periodo = int(df["id_paciente"].nunique())
        else:
            pacientes_periodo = 0

        ticket = round(ingresos / pacientes_periodo, 2) if pacientes_periodo > 0 else 0.0

        # ==============================
        # TOP tratamientos
        # ==============================
        if not df.empty:
            top_trat_ing = (
                df.groupby("tratamiento")["subtotal"]
                .sum()
                .sort_values(ascending=False)
                .head(5)
                .to_dict()
            )
            top_trat_vol = (
                df["tratamiento"]
                .value_counts()
                .head(5)
                .to_dict()
            )
        else:
            top_trat_ing = {}
            top_trat_vol = {}

        # ==============================
        # TOP doctores
        # ==============================
        if "doc_nombre" not in df.columns:
            df["doc_nombre"] = ""
        df["doc_nombre"] = df["doc_nombre"].fillna("").astype(str)

        if not df.empty:
            top_doc_ing = (
                df.groupby("doc_nombre")["subtotal"]
                .sum()
                .sort_values(ascending=False)
                .head(5)
                .to_dict()
            )
            top_doc_vol = (
                df["doc_nombre"]
                .value_counts()
                .head(5)
                .to_dict()
            )
        else:
            top_doc_ing = {}
            top_doc_vol = {}

        # ==============================
        # Comisiones doctor / promotor
        # ==============================

        def _normalize_name_local(s):
            if not s:
                return ""
            s = str(s).strip().lower()
            import unicodedata
            s = "".join(ch for ch in unicodedata.normalize("NFKD", s)
                        if not unicodedata.combining(ch))
            return " ".join(s.split())

        doctors_map_norm = {
            _normalize_name_local(k): float(v)
            for k, v in (self.cfg.doctors_percent or {}).items()
            if k is not None
        }

        personal_df = self.repo.dfs.get("Personal", pd.DataFrame()).copy()
        id_to_name_norm = {}
        if not personal_df.empty:
            for _, r in personal_df.iterrows():
                pid = str(r.get("id_personal", "")).strip()
                if pid:
                    id_to_name_norm[pid] = _normalize_name_local(r.get("nombre", ""))

        def pct_for_row(row):
            # 1) por id_doctor
            id_doc = str(row.get("id_doctor", "") or "").strip()
            if id_doc and id_doc in id_to_name_norm:
                return doctors_map_norm.get(id_to_name_norm[id_doc], 0.0)

            # 2) por nombre
            name = _normalize_name_local(row.get("doc_nombre", "") or "")
            if name in doctors_map_norm:
                return doctors_map_norm[name]

            # 3) coincidencia parcial
            for k in doctors_map_norm:
                if name and (name.startswith(k) or k.startswith(name) or k in name or name in k):
                    return doctors_map_norm[k]
            return 0.0

        df["pct_doc"] = df.apply(pct_for_row, axis=1)
        df["comision_doctor"] = (df["subtotal"] * df["pct_doc"]).round(2)

        if "pro_nombre" not in df.columns:
            df["pro_nombre"] = ""
        df["pro_nombre"] = df["pro_nombre"].fillna("").astype(str)

        promoter_pct = float(self.cfg.promoter_percent or 0.0)

        def _promo_calc(row):
            prom = row.get("pro_nombre", "")
            if str(prom).strip() != "":
                return round(float(row["subtotal"]) * promoter_pct, 2)
            return 0.0

        df["comision_promotor"] = df.apply(_promo_calc, axis=1)

        if "costo_tecnico" not in df.columns:
            df["costo_tecnico"] = 0.0
        df["costo_tecnico"] = pd.to_numeric(df["costo_tecnico"], errors="coerce").fillna(0.0)

        tot_cd = float(df["comision_doctor"].sum())
        tot_cp = float(df["comision_promotor"].sum())
        tot_ct = float(df["costo_tecnico"].sum())

        utilidad = round(ingresos - tot_cd - tot_cp - tot_ct, 2)

        # ==============================
        # Retención
        # ==============================
        all_cons = self.repo.dfs["Consultas"][["id_consulta", "id_paciente", "fecha"]].copy()
        all_cons["fecha"] = pd.to_datetime(all_cons["fecha"], errors="coerce")

        if "id_paciente" in df.columns:
            period_pac = set(df["id_paciente"].unique().tolist())
        else:
            period_pac = set()

        had_before = 0
        for pid in period_pac:
            prev = all_cons[
                (all_cons["id_paciente"] == pid) &
                (all_cons["fecha"] < pd.Timestamp(desde))
            ]
            if not prev.empty:
                had_before += 1

        retencion = round(had_before / len(period_pac), 4) if period_pac else 0.0

        # ==============================
        # Resultado
        # ==============================
        result = {
            "fecha_inicio": desde.isoformat(),
            "fecha_fin": hasta.isoformat(),
            "ingresos_periodo": round(ingresos, 2),
            "pacientes_unicos": pacientes_periodo,
            "ticket_promedio": ticket,
            "retencion": retencion,
            "utilidad_periodo": utilidad,
            "top_tratamientos_ingresos": top_trat_ing,
            "top_tratamientos_volumen": top_trat_vol,
            "top_dentistas_ingresos": top_doc_ing,
            "top_dentistas_volumen": top_doc_vol,
            "total_comisiones_doctores": tot_cd,
            "total_comisiones_promotores": tot_cp,
            "total_costo_tecnico": tot_ct,
        }

        # Snapshot KPIs_Export
        kdf = self.repo.dfs.get(
            "KPIs_Export",
            pd.DataFrame(columns=SHEETS.get("KPIs_Export", []))
        )

        snap = {
            "fecha_inicio": result["fecha_inicio"],
            "fecha_fin": result["fecha_fin"],
            "ingresos_periodo": ingresos,
            "pacientes_unicos": pacientes_periodo,
            "ticket_promedio": ticket,
            "top_tratamientos_json": json.dumps(top_trat_ing, ensure_ascii=False),
            "top_dentistas_json": json.dumps(top_doc_ing, ensure_ascii=False),
            "retencion": retencion,
            "utilidad_periodo": utilidad,
        }

        self.repo.dfs["KPIs_Export"] = pd.concat(
            [kdf, pd.DataFrame([snap])],
            ignore_index=True
        )
        self.repo.save()

        return result



    
    def ventas_totales_por_periodo(self, desde: date, hasta: date, periodo: str = "dia") -> pd.DataFrame:
        """
        Devuelve un DataFrame con columnas ['fecha','ventas'] agrupado por 'dia' | 'semana' | 'mes'
        - 'desde' y 'hasta' son date (inclusive)
        - 'periodo' puede ser "dia", "semana" o "mes"
        """
        rep = ReporteService(self.repo, self.cfg)
        df = rep.filtrar(desde, hasta, None, None, None, None, None).copy()
        # Asegurar columna fecha como datetime y subtotal numérico
        if "fecha" in df.columns:
            df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        else:
            # si no existe, intentar traer desde Consultas o devolver vacío
            return pd.DataFrame(columns=["fecha", "ventas"])

        df["subtotal"] = pd.to_numeric(df.get("subtotal", 0.0), errors="coerce").fillna(0.0)

        if periodo == "semana":
            # periodo representado por inicio de semana (lunes)
            df["periodo"] = df["fecha"].dt.to_period("W").apply(lambda r: r.start_time)
        elif periodo == "mes":
            df["periodo"] = df["fecha"].dt.to_period("M").apply(lambda r: r.start_time)
        else:
            # por día (fecha normalizada a medianoche)
            df["periodo"] = df["fecha"].dt.normalize()

        ventas = df.groupby("periodo", dropna=False)["subtotal"].sum().reset_index().rename(columns={"periodo": "fecha", "subtotal": "ventas"})
        # Asegurar tipo datetime en 'fecha'
        ventas["fecha"] = pd.to_datetime(ventas["fecha"], errors="coerce")
        # Ordenar ascendente
        ventas = ventas.sort_values("fecha").reset_index(drop=True)
        return ventas



# ---------------------------
# CLI (opcional)
# ---------------------------

def build_parser():
    p = argparse.ArgumentParser(description="Backend consultorio_items")
    p.add_argument("--config", required=True)
    sp = p.add_subparsers(dest="cmd")
    sp.add_parser("init")
    return p

def main(argv=None):
    args = build_parser().parse_args(argv)
    cfg = AppConfig.from_json(Path(args.config))
    if args.cmd == "init":
        _init_template(cfg.excel_path, cfg.overwrite_template)
        print("Plantilla creada:", cfg.excel_path)

# -------------------------------------------------------
# Exportar PDF del corte (añadir al final de consultorio_items.py)
# -------------------------------------------------------
def export_pdf_corte(cfg: AppConfig, id_corte: str, resumen: dict, detalle_df) -> str:
    """
    Genera un PDF simple con el resumen del corte y el detalle (DataFrame).
    Retorna la ruta como string. Si reportlab no está disponible, guarda un CSV alternativo.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import cm
    except Exception as e:
        # Si falta reportlab, exportamos a CSV como fallback (para no romper la GUI)
        try:
            out_dir = Path(cfg.pdf_output_dir) if cfg and cfg.pdf_output_dir else cfg.excel_path.parent
        except Exception:
            out_dir = Path(".")
        out_dir.mkdir(parents=True, exist_ok=True)
        csv_path = out_dir / f"corte_{id_corte}_detalle.csv"
        try:
            detalle_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        except Exception:
            # si ni siquiera podemos CSV, devolvemos string vacío
            return str(csv_path)
        return str(csv_path)

    out_dir = Path(cfg.pdf_output_dir) if cfg and cfg.pdf_output_dir else cfg.excel_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"corte_{id_corte}.pdf"

    # --- Start PDF ---
    c = canvas.Canvas(str(out_path), pagesize=A4)
    width, height = A4
    margin_x = 2 * cm
    y = height - 2 * cm

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawString(margin_x, y, f"Corte: {id_corte}")
    y -= 0.6 * cm
    c.setFont("Helvetica", 10)
    c.drawString(margin_x, y, f"Fecha de corte: {resumen.get('fecha_corte', '')}")
    y -= 0.5 * cm
    c.drawString(margin_x, y, f"Generado: {resumen.get('creado_en', '')}")
    y -= 0.8 * cm

    # Resumen numérico
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin_x, y, "Resumen:")
    y -= 0.5 * cm
    c.setFont("Helvetica", 10)
    resumen_items = [
        ("Total ingresos", resumen.get("total_ingresos", 0)),
        ("Total costo técnico", resumen.get("total_costo_tecnico", 0)),
        ("Total comisiones doctores", resumen.get("total_comisiones_doctores", 0)),
        ("Total comisiones promotores", resumen.get("total_comisiones_promotores", 0)),
        ("Total utilidad", resumen.get("total_utilidad", 0)),
    ]
    for label, val in resumen_items:
        c.drawString(margin_x + 0.4*cm, y, f"{label}: {float(val):.2f}")
        y -= 0.4 * cm

    y -= 0.4 * cm

    # Resumen por persona (doctor / promotor / tecnico) calculado desde detalle_df si es posible
    def draw_person_summary(title, mapping, x_off=margin_x):
        nonlocal y
        if y < 4 * cm:
            c.showPage()
            y = height - 2 * cm
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_off, y, title)
        y -= 0.45 * cm
        c.setFont("Helvetica", 9)
        if not mapping:
            c.drawString(x_off + 0.4*cm, y, "—")
            y -= 0.35 * cm
            return
        for k, v in mapping.items():
            lbl = str(k) if (k and str(k).strip()) else "—"
            c.drawString(x_off + 0.4*cm, y, f"{lbl}: {float(v):.2f}")
            y -= 0.35 * cm
            if y < 4 * cm:
                c.showPage()
                y = height - 2 * cm

    # try to compute mappings from detalle_df
    try:
        if detalle_df is not None and hasattr(detalle_df, "groupby"):
            # doctor
            if "doc_nombre" in detalle_df.columns and "comision_doctor" in detalle_df.columns:
                doc_map = detalle_df.groupby("doc_nombre")["comision_doctor"].sum().round(2).to_dict()
            else:
                doc_map = {}
            if "pro_nombre" in detalle_df.columns and "comision_promotor" in detalle_df.columns:
                pro_map = detalle_df.groupby("pro_nombre")["comision_promotor"].sum().round(2).to_dict()
            else:
                pro_map = {}
            if "tec_nombre" in detalle_df.columns and "costo_tecnico" in detalle_df.columns:
                tec_map = detalle_df.groupby("tec_nombre")["costo_tecnico"].sum().round(2).to_dict()
            else:
                tec_map = {}
        else:
            doc_map = pro_map = tec_map = {}
    except Exception:
        doc_map = pro_map = tec_map = {}

    draw_person_summary("Doctores (a pagar)", doc_map, margin_x)
    draw_person_summary("Promotores (a pagar)", pro_map, margin_x)
    draw_person_summary("Técnicos (costo técnico)", tec_map, margin_x)

    # Detalle de items (tabla simplificada)
    """y -= 0.4 * cm
    if y < 6 * cm:
        c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 11); c.drawString(margin_x, y, "Detalle (items)"); y -= 0.6 * cm
    c.setFont("Helvetica", 9)

    # Columns to show (truncate long text)
    cols = ["fecha", "paciente", "tratamiento", "descripcion_item", "cantidad", "precio_unitario", "subtotal", "doc_nombre", "pro_nombre", "tec_nombre"]
    col_widths = [2.2*cm, 3.4*cm, 3.2*cm, 4.0*cm, 1.0*cm, 2.0*cm, 2.0*cm, 2.6*cm, 2.6*cm, 2.6*cm]
    header_x = margin_x
    for i, col in enumerate(cols):
        c.drawString(header_x, y, col[:15])
        header_x += col_widths[i]
    y -= 0.45 * cm

    # rows
    max_rows = 40  # evitar páginas infinitas, se puede paginate
    rows = 0
    if detalle_df is not None:
        for _, row in (detalle_df.fillna("").iterrows()):
            if rows >= max_rows:
                # indica que hay más
                c.drawString(margin_x, y, "... (más filas en el archivo CSV/PDF completo)")
                y -= 0.4*cm
                break
            if y < 3.5 * cm:
                c.showPage(); y = height - 2*cm
            line_x = margin_x
            for i, col in enumerate(cols):
                val = row.get(col, "")
                if isinstance(val, (float, int)):
                    txt = f"{float(val):.2f}"
                else:
                    txt = str(val)
                txt = txt.replace("\n"," ")[:int(col_widths[i]/0.18)]  # truncate approx
                c.drawString(line_x, y, txt)
                line_x += col_widths[i]
            y -= 0.35 * cm
            rows += 1"""

    c.showPage()
    c.save()
    return str(out_path)

# ============================================================
# EXPORTAR PDF DE KPIs
# ============================================================
def export_pdf_kpis(cfg, kpis_dict):
    """
    Genera un PDF con los KPIs calculados.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    import os
    from datetime import datetime

    # Crear carpeta "reports" al lado del Excel si no existe
    out_dir = cfg.excel_path.parent / "reports"
    out_dir.mkdir(exist_ok=True)

    fecha_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
    pdf_path = out_dir / f"KPI_{fecha_str}.pdf"

    doc = SimpleDocTemplate(str(pdf_path), pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>Reporte de KPIs – Consultorio</b>", styles["Title"]))
    story.append(Spacer(1, 0.2 * inch))

    resumen = [
        ["Ingresos", f"${kpis_dict.get('ingresos_periodo', 0):,.2f}"],
        ["Utilidad (consultorio)", f"${kpis_dict.get('utilidad', 0):,.2f}"],
        ["Pacientes únicos", str(kpis_dict.get("pacientes_unicos", 0))],
        ["Ticket promedio", f"${kpis_dict.get('ticket_promedio', 0):,.2f}"],
        ["Retención", f"{kpis_dict.get('retencion', 0)*100:.1f}%"]
    ]

    table = Table(resumen, colWidths=[2.5*inch, 2*inch])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")
    ]))
    story.append(table)
    story.append(Spacer(1, 0.3 * inch))

    # Top tratamientos
    story.append(Paragraph("<b>Top tratamientos (ingresos)</b>", styles["Heading3"]))
    for k, v in (kpis_dict.get("top_tratamientos_ingresos", {}) or {}).items():
        story.append(Paragraph(f"{k}: ${v:,.2f}", styles["Normal"]))
    story.append(Spacer(1, 0.2 * inch))

    story.append(Paragraph("<b>Top dentistas (ingresos)</b>", styles["Heading3"]))
    for k, v in (kpis_dict.get("top_dentistas_ingresos", {}) or {}).items():
        story.append(Paragraph(f"{k}: ${v:,.2f}", styles["Normal"]))
    story.append(Spacer(1, 0.2 * inch))

    doc.build(story)
    return pdf_path

if __name__ == "__main__":
    main()

