#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI con soporte para items por consulta.
Archivo: consultorio_gui_items.py
Requiere: consultorio_items.py en la misma carpeta.
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
from pathlib import Path
from datetime import date
from datetime import datetime
from datetime import timedelta
from tkcalendar import DateEntry
import json
import matplotlib
#matplotlib.use("TkAgg")   # para evitar problemas en entornos sin display; cambiar si necesitas backends interactivos
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import hashlib
from licencia_local import evaluar_licencia

# Import backend
try:
    from consultorio_items import AppConfig, Repo, RegistroService, ReporteService, CorteService, KPIService, ComisionService, _parse_date
except Exception as e:
    raise SystemExit("No puedo importar consultorio_items.py: " + str(e))

APP_TITLE = "APP CONSULTORIO – GUI (Items)"
DEFAULT_CONFIG_PATH = Path("./config.json").resolve()

class ConsultorioGUI(tk.Tk):
    def __init__(self, dias_licencia_restantes: int | None = None):
        super().__init__()

        self.dias_licencia_restantes = dias_licencia_restantes

        titulo = APP_TITLE
        if self.dias_licencia_restantes is not None:
            titulo = f"{APP_TITLE} (Prueba – {self.dias_licencia_restantes} días restantes)"

        self.title(titulo)
        self.geometry("1200x780")
        self.minsize(1000,700)

        self.config_path = tk.StringVar(value=str(DEFAULT_CONFIG_PATH))
        self.cfg = None
        self.repo = None
        self._comision_service = None

        self._build_menu()

        # === Notebook ===
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        # === Crear pestañas ===
        self.tab_registro      = ttk.Frame(self.nb)
        self.tab_espera        = ttk.Frame(self.nb)
        self.tab_agenda        = ttk.Frame(self.nb)
        self.tab_historial     = ttk.Frame(self.nb)
        self.tab_reportes      = ttk.Frame(self.nb)
        self.tab_cortes        = ttk.Frame(self.nb)
        self.tab_kpis          = ttk.Frame(self.nb)
        self.tab_config        = ttk.Frame(self.nb)
        self.tab_comisiones    = ttk.Frame(self.nb)   # ← NUEVA pestaña oficial

        # === Agregar pestañas al Notebook ===
        self.nb.add(self.tab_registro, text="Registro")
        self.nb.add(self.tab_espera, text="Consultas en espera")
        self.nb.add(self.tab_agenda, text="Agenda")
        self.nb.add(self.tab_reportes, text="Reportes")
        self.nb.add(self.tab_cortes, text="Cortes")
        self.nb.add(self.tab_kpis, text="KPIs")
        self.nb.add(self.tab_config, text="Config")
        self.nb.add(self.tab_comisiones, text="Comisiones")  # ← Aquí aparece en UI

        # === Construir interfaces de cada pestaña ===
        self._build_config_tab()
        self._build_registro_tab()
        self._build_historial_tab()
        self._build_reportes_tab()
        self._build_cortes_tab()
        self._build_kpis_tab()
        self._build_espera_tab()
        self._build_agenda_tab()
        self._build_comisiones_tab()   # ← IMPORTANTE: construir pestaña comisiones

        # === Cargar config inicial ===
        self._load_config_silent()

        # Forzar foco
        self.after(300, lambda: self.focus_force())


    def _build_menu(self):
        menubar = tk.Menu(self); filemenu = tk.Menu(menubar, tearoff=0); filemenu.add_command(label="Abrir config.json…", command=self._choose_config); filemenu.add_separator(); filemenu.add_command(label="Salir", command=self.destroy); menubar.add_cascade(label="Archivo", menu=filemenu); self.config(menu=menubar)
############################
    def _build_config_tab(self):
        frm = self.tab_config
        # encabezado similar al anterior
        row1 = ttk.Frame(frm); row1.pack(fill="x", pady=10, padx=12)
        ttk.Label(row1, text="Ruta de config.json:").pack(side="left")
        ttk.Entry(row1, textvariable=self.config_path, width=70).pack(side="left", padx=6)
        ttk.Button(row1, text="Examinar…", command=self._choose_config).pack(side="left")
        ttk.Button(row1, text="Cargar", command=self._load_config).pack(side="left", padx=6)

        self.lbl_excel_path = ttk.Label(frm, text="Excel: —"); self.lbl_excel_path.pack(anchor="w", padx=12)
        row2 = ttk.Frame(frm); row2.pack(fill="x", pady=6, padx=12)
        ttk.Button(row2, text="Crear plantilla (init)", command=self._cmd_init_template).pack(side="left")
        ttk.Button(row2, text="Abrir carpeta del Excel", command=self._open_excel_folder).pack(side="left", padx=10)
        self.lbl_status = ttk.Label(frm, text="Estado: sin configuración cargada", foreground="#555"); self.lbl_status.pack(anchor="w", padx=12, pady=10)

        # bloqueo admin
        lock_row = ttk.Frame(frm); lock_row.pack(fill="x", padx=12, pady=6)
        self._cfg_unlocked = False
        self._cfg_current_user = None
        ttk.Label(lock_row, text="Config segura: ").pack(side="left")
        self._btn_unlock = ttk.Button(lock_row, text="Desbloquear (admin)", command=self._prompt_admin_password)
        self._btn_unlock.pack(side="left", padx=6)
        ttk.Button(lock_row, text="Cambiar contraseña admin", command=self._set_admin_password).pack(side="left", padx=6)

        # Panel de personal y controles
        pers_frame = ttk.LabelFrame(frm, text="Personal (comisiones)")
        pers_frame.pack(fill="both", expand=True, padx=12, pady=8)

        # Treeview de personal
        cols = ("id_personal","nombre","rol","activo","pct_comision","modificado_en")
        self.tree_personal = ttk.Treeview(pers_frame, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree_personal.heading(c, text=c)
            self.tree_personal.column(c, width=120 if c not in ("nombre","rol") else 180, anchor="center")
        self.tree_personal.pack(fill="both", expand=True, padx=6, pady=6)

        btns = ttk.Frame(pers_frame); btns.pack(fill="x", padx=6, pady=4)
        ttk.Button(btns, text="Agregar personal", command=self._add_personal).pack(side="left")
        ttk.Button(btns, text="Editar % seleccionado", command=self._edit_personal).pack(side="left", padx=6)
        ttk.Button(btns, text="Desactivar seleccionado", command=self._delete_personal).pack(side="left", padx=6)
        ttk.Button(btns, text="Ver historial", command=self._view_personal_history).pack(side="left", padx=6)
        ttk.Button(btns, text="Eliminar seleccionado", command=self._eliminar_personal_sel).pack(side="left", padx=6)



        # Info de ayuda
        self.lbl_config_note = ttk.Label(frm, text="Autentícate para editar personal y porcentajes.", foreground="#666")
        self.lbl_config_note.pack(anchor="w", padx=12, pady=(4,10))

        # carga inicial (si repo/cfg ya cargada)
        try:
            if self.repo:
                self._load_personal_table()
        except Exception:
            pass

    def _hash_password(self, pwd: str) -> str:
        return hashlib.sha256(pwd.encode("utf-8")).hexdigest()
    
    def _prompt_admin_password(self):
        """Solicita contraseña admin; si cfg tiene hash lo valida; si no, permite crear una nueva (confirmando)."""
        if not self.cfg:
            messagebox.showwarning("Config","Carga config.json primero.")
            return
    # si no hay hash en cfg, pedir crear una
        if not getattr(self.cfg, "admin_password_hash", None):
            ok = messagebox.askyesno("Contraseña admin", "No hay contraseña configurada. ¿Deseas crear una ahora?")
            if not ok:
                return
            # pedir y confirmar
            pwd1 = tk.simpledialog.askstring("Nueva contraseña", "Contraseña nueva (se almacenará hash)", show="*")
            if not pwd1:
                return
            pwd2 = tk.simpledialog.askstring("Confirmar", "Confirma la contraseña", show="*")
            if pwd1 != pwd2:
                messagebox.showerror("Contraseña", "No coinciden.")
                return
            # guardar en config.json (modificar el archivo)
            try:
                p = Path(self.config_path.get())
                cfg_raw = json.loads(p.read_text(encoding="utf-8"))
                cfg_raw["admin_password_hash"] = self._hash_password(pwd1)
                p.write_text(json.dumps(cfg_raw, ensure_ascii=False, indent=2), encoding="utf-8")
                messagebox.showinfo("Contraseña", "Contraseña guardada en config.json")
                # recargar cfg
                self._load_config()
            except Exception as e:
                messagebox.showerror("Contraseña", f"No pude guardar la contraseña:\n{e}")
            return

        # si ya hay hash, pedir password y validar
        pwd = tk.simpledialog.askstring("Contraseña admin", "Introduce la contraseña de administrador", show="*")
        if not pwd:
            return
        if self._hash_password(pwd) == getattr(self.cfg, "admin_password_hash", ""):
            self._cfg_unlocked = True
            self._cfg_current_user = "admin"
            messagebox.showinfo("Config", "Autenticación correcta. Puedes editar ahora.")
            self._btn_unlock.config(text="Bloqueado (desconectar)", command=self._lock_config)
            self._load_personal_table()
        else:
            messagebox.showerror("Config", "Contraseña incorrecta.")

    def _lock_config(self):
        self._cfg_unlocked = False
        self._cfg_current_user = None
        messagebox.showinfo("Config", "Configuración bloqueada.")
        self._btn_unlock.config(text="Desbloquear (admin)", command=self._prompt_admin_password)
        # limpiar selección
        try:
            self.tree_personal.selection_remove(self.tree_personal.selection())
        except Exception:
            pass

    def _set_admin_password(self):
        """Permite cambiar contraseña: pide actual, luego nueva."""
        if not self.cfg:
            messagebox.showwarning("Config","Carga config.json primero.")
            return
        # si no hay password, redirigir a crear (llama _prompt_admin_password)
        if not getattr(self.cfg, "admin_password_hash", None):
            return self._prompt_admin_password()
        # pedir actual
        cur = tk.simpledialog.askstring("Contraseña actual", "Introduce la contraseña actual", show="*")
        if not cur:
            return
        if self._hash_password(cur) != getattr(self.cfg, "admin_password_hash", ""):
            messagebox.showerror("Contraseña", "Contraseña actual incorrecta.")
            return
        # pedir nueva
        new1 = tk.simpledialog.askstring("Nueva contraseña", "Nueva contraseña", show="*")
        if not new1:
            return
        new2 = tk.simpledialog.askstring("Confirmar", "Confirma la nueva contraseña", show="*")
        if new1 != new2:
            messagebox.showerror("Contraseña", "No coinciden.")
            return
        # escribir en config.json
        try:
            p = Path(self.config_path.get())
            cfg_raw = json.loads(p.read_text(encoding="utf-8"))
            cfg_raw["admin_password_hash"] = self._hash_password(new1)
            p.write_text(json.dumps(cfg_raw, ensure_ascii=False, indent=2), encoding="utf-8")
            messagebox.showinfo("Contraseña", "Contraseña actualizada.")
            self._load_config()
        except Exception as e:
            messagebox.showerror("Contraseña", f"No pude actualizar contraseña:\n{e}")

# ---------- Gestión de personal (UI) ----------
    def _load_personal_table(self):
        """Carga tabla de Personal desde repo y la muestra en tree_personal."""
        if not self._ensure_repo():
            return
        df = self.repo.list_personal()
        # poblar tree
        self.tree_personal.delete(*self.tree_personal.get_children())
        for _, r in df.iterrows():
            vals = (r.get("id_personal",""), r.get("nombre",""), r.get("rol",""), int(r.get("activo",1)), f"{float(r.get('pct_comision',0.0)):.2f}", r.get("modificado_en",""))
            self.tree_personal.insert("", "end", values=vals)

    def _add_personal(self):
        if not self._cfg_unlocked:
            messagebox.showwarning("Config","Autentícate como admin para agregar personal.")
            return
        # pedir campos
        name = tk.simpledialog.askstring("Agregar personal", "Nombre completo")
        if not name: return
        rol = tk.simpledialog.askstring("Rol", "Rol (doctor/promotor/otro)", initialvalue="doctor")
        if not rol: return
        pct = tk.simpledialog.askfloat("Porcentaje", "Porcentaje de comisión (ej: 0.20 = 20%)", minvalue=0.0, maxvalue=1.0)
        if pct is None: return
        try:
            pid = self.repo.add_personal(name, rol, float(pct), usuario=self._cfg_current_user or "admin")
            messagebox.showinfo("Personal", f"Personal agregado: {pid}")
            self._load_personal_table()
        except Exception as e:
            messagebox.showerror("Personal", f"Error al agregar:\n{e}")

    def _edit_personal(self):
        if not self._cfg_unlocked:
            messagebox.showwarning("Config","Autentícate como admin para editar.")
            return
        sel = self.tree_personal.selection()
        if not sel:
            messagebox.showwarning("Personal","Selecciona una fila.")
            return
        iid = sel[0]
        vals = self.tree_personal.item(iid)["values"]
        idp = vals[0]
        nombre = vals[1]
        old_pct = float(vals[4])
        new_pct = tk.simpledialog.askfloat("Editar porcentaje", f"{nombre}\nPorcentaje actual: {old_pct:.2f}\nNuevo porcentaje (0-1)", minvalue=0.0, maxvalue=1.0)
        if new_pct is None:
            return
        ok = self.repo.update_personal_pct(idp, float(new_pct), usuario=self._cfg_current_user or "admin")
        if ok:
            messagebox.showinfo("Personal", "Porcentaje actualizado.")
            self._load_personal_table()
        else:
            messagebox.showerror("Personal", "No pude actualizar.")

    def _delete_personal(self):
        if not self._cfg_unlocked:
            messagebox.showwarning("Config","Autentícate como admin para desactivar.")
            return
        sel = self.tree_personal.selection()
        if not sel:
            messagebox.showwarning("Personal","Selecciona una fila.")
            return
        iid = sel[0]
        vals = self.tree_personal.item(iid)["values"]
        idp = vals[0]
        nombre = vals[1]
        if not messagebox.askyesno("Desactivar", f"Desactivar {nombre}?"):
            return
        ok = self.repo.deactivate_personal(idp, usuario=self._cfg_current_user or "admin")
        if ok:
            messagebox.showinfo("Personal", "Personal desactivado.")
            self._load_personal_table()
        else:
            messagebox.showerror("Personal", "No pude desactivar.")

    def _view_personal_history(self):
        if not self._ensure_repo():
            return

        try:
            hist = self.repo.get_personal_history()

            if hist.empty:
                messagebox.showinfo("Historial", "No hay registros.")
                return

            # Quitar NaN
            hist = hist.fillna("")

            # Crear ventana
            win = tk.Toplevel(self)
            win.title("Historial de personal")
            win.geometry("1000x600")

            frm = ttk.Frame(win)
            frm.pack(fill="both", expand=True)

            # 🔹 OJO: Ya NO va la columna "usuario"
            cols = ("fecha", "accion", "nombre", "rol", "valor_anterior", "valor_nuevo")

            tree = ttk.Treeview(frm, columns=cols, show="headings", height=20)

            tree.heading("fecha", text="Fecha")
            tree.heading("accion", text="Acción")
            tree.heading("nombre", text="Nombre")
            tree.heading("rol", text="Rol")
            tree.heading("valor_anterior", text="Valor anterior")
            tree.heading("valor_nuevo", text="Valor nuevo")

            tree.column("fecha", width=150, anchor="center")
            tree.column("accion", width=130, anchor="center")
            tree.column("nombre", width=200, anchor="center")
            tree.column("rol", width=120, anchor="center")
            tree.column("valor_anterior", width=150, anchor="center")
            tree.column("valor_nuevo", width=150, anchor="center")

            vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
            tree.configure(yscroll=vsb.set, xscroll=hsb.set)

            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")

            frm.rowconfigure(0, weight=1)
            frm.columnconfigure(0, weight=1)

            # Preparar filas PARA LAS NUEVAS COLUMNAS
            rows = []
            for _, r in hist.iterrows():
                rows.append([
                    str(r.get("fecha", "")),
                    r.get("accion", ""),
                    r.get("nombre", ""),
                    r.get("rol", ""),
                    r.get("valor_anterior", ""),
                    r.get("valor_nuevo", "")
                ])

            # Zebra stripes con tus helpers
            self._zebra_insert_all(tree, rows)

        except Exception as e:
            messagebox.showerror("Historial", f"Error: {e}")


    def _eliminar_personal_sel(self):
        if not self._ensure_repo():
            return

        sel = self.tree_personal.selection()
        if not sel:
            messagebox.showwarning("Personal", "Selecciona un registro.")
            return

        iid = sel[0]
        vals = self.tree_personal.item(iid)["values"]

        # Asumiendo orden de columnas: id_personal, nombre, rol, ...
        id_personal = str(vals[0])
        nombre = str(vals[1])

        if not messagebox.askyesno("Eliminar",
                                   f"¿Eliminar definitivamente a '{nombre}' del catálogo de personal?"):
            return

        usuario = self._cfg_current_user or "admin"
        ok = self.repo.eliminar_personal(id_personal, usuario=usuario)

        if ok:
            self._load_personal_table()
        else:
            messagebox.showerror("Personal", "No se pudo eliminar el registro seleccionado.")

  
    def _refrescar_personal(self):
        if not self._ensure_repo():
            return

        # Leer directamente del DataFrame
        df = self.repo.dfs.get("Personal", pd.DataFrame())

        # Mostrar solo activos
        df = self._filtrar_personal_activo(df)

        rows = []
        for _, r in df.iterrows():
            vals = [
                r.get("id_personal",""),
                r.get("nombre",""),
                r.get("rol",""),
                r.get("activo",""),
                r.get("pct_comision",""),
                r.get("modificado_en","")
            ]
            rows.append(vals)

        self._zebra_insert_all(self.tree_personal, rows)


    ####################################
    def _choose_config(self):
        p = filedialog.askopenfilename(title="Selecciona config.json", filetypes=[("JSON","*.json"),("Todos","*.*")])
        if p: self.config_path.set(p)

    def _load_config_silent(self):
        p = Path(self.config_path.get())
        if p.exists():
            try: self._load_config()
            except Exception: pass

    def _load_config(self):
        try:
            cfg = AppConfig.from_json(Path(self.config_path.get()))
            self.cfg = cfg
            if cfg.excel_path.exists():
                self.repo = Repo(cfg)
                status = "OK – Excel cargado"
                # Al cargar la configuración, refrescar el personal para que aparezca
                # inmediatamente en los combos de Comisiones y en la tabla de Config.
                self._cargar_personal_comisiones()
                self._load_personal_table()
            else:
                self.repo = None
                status = "Excel NO encontrado. Crea plantilla"
            self.lbl_excel_path.config(text=f"Excel: {cfg.excel_path}")
            self.lbl_status.config(text=f"Estado: {status}")
            messagebox.showinfo("Config", "Configuración cargada.")
        except Exception as e:
            messagebox.showerror("Config", f"No pude cargar config.json\n{e}")

    def _cmd_init_template(self):
        if not self.cfg: messagebox.showwarning("Init","Carga config.json primero"); return
        try:
            from consultorio_items import _init_template
            _init_template(self.cfg.excel_path, self.cfg.overwrite_template)
            self.repo = Repo(self.cfg)
            self.lbl_status.config(text="Estado: OK – Plantilla creada y cargada")
            messagebox.showinfo("Init","Plantilla creada.")
        except Exception as e:
            messagebox.showerror("Init", f"Error: {e}")

    def _open_excel_folder(self):
        if not self.cfg: return
        folder = self.cfg.excel_path.parent
        import os, platform, subprocess
        try:
            if platform.system()=="Windows": os.startfile(str(folder))
            elif platform.system()=="Darwin": subprocess.run(["open", str(folder)])
            else: subprocess.run(["xdg-open", str(folder)])
        except: pass

    def _mk_date_entry(self, parent, var=None, default=None):
        de = DateEntry(parent, width=12, date_pattern="yyyy-mm-dd", locale="es_MX", showweeknumbers=False)

        if default:
            try:
                de.set_date(default)
            except:
                pass

        if var is not None:
            def _sync_var(*_):
                d = de.get_date()
                var.set(d.isoformat())
            de.bind("<<DateEntrySelected>>", _sync_var)
            _sync_var()

        # -----------------------------------------------------------
        # 🔧 FIX PARA QUE LA VENTANA NO SE CIERRE AL CAMBIAR MES/AÑO
        # -----------------------------------------------------------
        def apply_fix():
            try:
                topcal = de._top_cal
            except:
                # el popup aún no existe -> volver a intentar
                de.after(20, apply_fix)
                return

            # una vez existe, interceptamos el FocusOut
            try:
                topcal.unbind("<FocusOut>")
                topcal.bind("<FocusOut>", lambda e: "break")
            except:
                pass

        # cuando se abre el calendario: intentar aplicar fix
        def on_open(evt):
            de.after(20, apply_fix)

        # enganchar apertura del calendario
        try:
            de.bind("<<DateEntryOpen>>", on_open)
        except:
            pass
        #----------------------------------------------------

        return de
 
 ###############################cambio de registri de fecha########################

    def _mk_date_selector(self, parent, var=None, default=None):
        import calendar
        from datetime import date

        frame = ttk.Frame(parent)

        # Valores iniciales
        today = default or date.today()
        year0 = today.year
        month0 = today.month
        day0 = today.day

        # Combobox Año
        years = [str(y) for y in range(year0 - 5, year0 + 6)]
        cb_year = ttk.Combobox(frame, values=years, width=6, state="readonly")
        cb_year.set(str(year0))
        cb_year.pack(side="left", padx=2)

        # Combobox Mes
        months = [f"{m:02d}" for m in range(1, 12+1)]
        cb_month = ttk.Combobox(frame, values=months, width=4, state="readonly")
        cb_month.set(f"{month0:02d}")
        cb_month.pack(side="left", padx=2)

        # Combobox Día (se actualizará según el mes/año)
        def update_days(*args):
            y = int(cb_year.get())
            m = int(cb_month.get())
            ndays = calendar.monthrange(y, m)[1]
            cb_day["values"] = [f"{d:02d}" for d in range(1, ndays+1)]
            # Ajustar si día previo ya no existe en este mes
            try:
                if int(cb_day.get()) > ndays:
                    cb_day.set(f"{ndays:02d}")
            except:
                cb_day.set("01")

            # actualizar variable final
            if var is not None:
                var.set(f"{cb_year.get()}-{cb_month.get()}-{cb_day.get()}")

        cb_month.bind("<<ComboboxSelected>>", update_days)
        cb_year.bind("<<ComboboxSelected>>", update_days)

        # Día
        cb_day = ttk.Combobox(frame, values=[], width=4, state="readonly")
        cb_day.pack(side="left", padx=2)

        # Inicializar días según default
        update_days()
        cb_day.set(f"{day0:02d}")
        cb_day.pack_configure()

        # Sincronizar cambios
        def apply_change(*args):
            if var is not None:
                var.set(f"{cb_year.get()}-{cb_month.get()}-{cb_day.get()}")

        cb_day.bind("<<ComboboxSelected>>", apply_change)

        # Set inicial
        apply_change()

        return frame
 
 
 ########################################## detalles en tabla zebta y scrollbars ##########################   
    def _apply_zebra(self, tree):
        """Pinta zebra stripes."""
        tree.tag_configure("oddrow", background="#F2F2F2")   # gris suave
        tree.tag_configure("evenrow", background="#FFFFFF")  # blanco

    def _zebra_insert_all(self, tree, values_list):
        """
        Recibe una lista de filas y las inserta con zebra stripes.
        values_list = [ [col1, col2, col3...], (...) ]
        """
        tree.delete(*tree.get_children())

        for i, vals in enumerate(values_list):
            tag = "oddrow" if i % 2 else "evenrow"
            tree.insert("", "end", values=vals, tags=(tag,))

    def _insert_zebra(self, tree, values, index):
        """Inserta fila con stripe correspondiente."""
        tag = "oddrow" if index % 2 else "evenrow"
        tree.insert("", "end", values=values, tags=(tag,))

    def _wrap_tree_with_scroll(self, parent, tree):
        """
        Envuelve un Treeview con scroll vertical y horizontal.
        - Crea un contenedor interno.
        - Crea un NUEVO Treeview dentro de ese contenedor (para que su master sea correcto).
        - Copia columnas/encabezados del tree original.
        - Usa grid dentro del contenedor, sin mezclar pack/grid en el mismo frame padre.
        """
        # 1) Contenedor interno que sí se packea en el parent
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        # 2) Extraer configuración del tree original
        try:
            cols = tree["columns"]
        except Exception:
            cols = ()

        show = tree["show"] if "show" in tree.keys() else "headings"
        # height viene como string a veces, lo normalizamos
        try:
            height = int(tree["height"])
        except Exception:
            height = 10

        # Guardar configuración de encabezados y columnas
        heading_cfg = {}
        column_cfg = {}
        for c in cols:
            try:
                heading_cfg[c] = tree.heading(c)
            except Exception:
                heading_cfg[c] = {}
            try:
                column_cfg[c] = tree.column(c)
            except Exception:
                column_cfg[c] = {}

        # 3) Destruir el tree original (estaba ligado al frame equivocado)
        try:
            tree.destroy()
        except Exception:
            pass

        # 4) Crear el NUEVO Treeview dentro del contenedor correcto
        new_tree = ttk.Treeview(container, columns=cols, show=show, height=height)

        # Restaurar encabezados y columnas
        for c in cols:
            cfg_h = heading_cfg.get(c, {})
            txt = cfg_h.get("text", c)
            new_tree.heading(c, text=txt)

            cfg_col = column_cfg.get(c, {})
            col_kwargs = {}
            for k in ("width", "anchor", "stretch", "minwidth"):
                if k in cfg_col:
                    col_kwargs[k] = cfg_col[k]
            if col_kwargs:
                new_tree.column(c, **col_kwargs)

        # 5) Scrollbars
        vsb = ttk.Scrollbar(container, orient="vertical", command=new_tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=new_tree.xview)
        new_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # 6) Layout con grid, pero SOLO dentro del contenedor
        new_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        return new_tree


############################################################################################
    # -------------------------
# Agenda diaria con intervalos de 15 minutos
# -------------------------

    def _build_agenda_tab(self):
        frm = self.tab_agenda

        # ======== Barra superior (Fecha + Botón) ========
        top = ttk.Frame(frm)
        top.pack(fill="x", padx=12, pady=10)

        ttk.Label(top, text="Fecha:").pack(side="left", padx=(0,6))

        # Variable de fecha
        self._agenda_fecha = tk.StringVar()

        # Selector de fecha (CORREGIDO: ahora se muestra)
        selector = self._mk_date_selector(top, var=self._agenda_fecha)
        selector.pack(side="left", padx=4)

        # Botón ver agenda
        ttk.Button(top, text="Ver agenda", command=self._agenda_refrescar).pack(side="left", padx=10)

        # ======== Contenedor central para tabla ========
        cont_slots = ttk.Frame(frm)
        cont_slots.pack(fill="both", expand=True, padx=12, pady=10)

        # Columnas
        cols = ("hora", "paciente", "doctor")

        # Treeview
        self.tree_slots = ttk.Treeview(cont_slots, columns=cols, show="headings", height=28)

        # Encabezados
        self.tree_slots.heading("hora", text="Hora")
        self.tree_slots.heading("paciente", text="Paciente")
        self.tree_slots.heading("doctor", text="Doctor")

        # Tamaños
        self.tree_slots.column("hora", width=90, anchor="center")
        self.tree_slots.column("paciente", width=180, anchor="center")
        self.tree_slots.column("doctor", width=150, anchor="center")

        # Wrap con scroll (sin mezclar pack/grid)
        self.tree_slots = self._wrap_tree_with_scroll(cont_slots, self.tree_slots)

        # Zebra stripes
        self._apply_zebra(self.tree_slots)

        # Evento de selección
        self.tree_slots.bind("<<TreeviewSelect>>", self._agenda_slot_select)

        # ======== Intervalos de horario ========
        self._AGENDA_INTERVALOS = [
            f"{h:02d}:{m:02d}"
            for h in range(8, 21)
            for m in (0,15,30,45)
        ]

    def _agenda_refrescar(self):
        if not self._ensure_repo():
            return

        fecha = self._agenda_fecha.get().strip()
        if not fecha:
            messagebox.showwarning("Agenda", "Selecciona una fecha.")
            return

        self.tree_slots.delete(*self.tree_slots.get_children())

        df = self._agenda_get_citas_del_dia(fecha)

        for hora in self._AGENDA_INTERVALOS:

            # Buscar cita que coincida con esta hora
            cita_row = None
            for _, r in df.iterrows():
                if r["hora_inicio"] <= hora < r["hora_fin"]:
                    cita_row = r.to_dict()   # <<< CONVERSIÓN CORRECTA
                    break

            if cita_row is not None:         # <<< NO MÁS ERROR
                pid = cita_row["id_paciente"]
                paciente = self._get_patient_name(pid)
                doctor = self._get_personal_name(cita_row.get("id_doctor"))
                self.tree_slots.insert(
                    "", "end",
                    values=(hora, paciente, doctor),
                    tags=("ocupado", cita_row["id_cita"])
    )
            else:
                # libre
                self.tree_slots.insert(
                    "", "end",
                    values=(hora, "", ""),
                    tags=("libre",)
                )
    
    def _agenda_get_citas_del_dia(self, fecha):
        df = self.repo.dfs["Citas"]

        if df.empty:
            return pd.DataFrame()

        # Normalizar fecha a string
        fecha = str(fecha)

        df = df[df["fecha"].astype(str) == fecha].copy()

        # Si no hay citas ese día, regresar df vacío inmediatamente
        if df.empty:
            return df

        # Crear columnas normalizadas
        df["hora_inicio"] = df["hora"].astype(str)

        # Asegurar que hora_fin exista
        if "hora_fin" not in df.columns:
            df["hora_fin"] = ""

        # Función auxiliar
        def fix_fin(val, hi):
            try:
                if isinstance(val, str) and val.strip():
                    return val
            except:
                pass
            t = datetime.strptime(hi, "%H:%M")
            t2 = t + timedelta(minutes=30)
            return t2.strftime("%H:%M")

        # Aplicar normalización solo si HAY filas
        df["hora_fin"] = df.apply(
            lambda r: fix_fin(r.get("hora_fin", ""), r["hora_inicio"]),
            axis=1
        )

        return df

    def _agenda_get_horas_ocupadas(self, fecha):
        """Devuelve lista de horas ocupadas como strings HH:MM."""
        df = self._agenda_get_citas_del_dia(fecha)
        ocupadas = set()

        for _, r in df.iterrows():
            hi = r["hora_inicio"]
            hf = r["hora_fin"]

            t1 = datetime.strptime(hi, "%H:%M")
            t2 = datetime.strptime(hf, "%H:%M")

            # añadir cada intervalo de 15 min dentro del rango
            t = t1
            while t < t2:
                ocupadas.add(t.strftime("%H:%M"))
                t += timedelta(minutes=15)

        return sorted(list(ocupadas))

    def _agenda_get_horas_libres(self, fecha):
        """Devuelve solo las horas libres."""
        ocupadas = set(self._agenda_get_horas_ocupadas(fecha))
        todas = set(self._AGENDA_INTERVALOS)
        libres = sorted(list(todas - ocupadas))
        return libres

    def _agenda_slot_select(self, event):
        sel = self.tree_slots.selection()
        if not sel:
            return

        item = sel[0]
        vals = self.tree_slots.item(item)["values"]
        tags = self.tree_slots.item(item)["tags"]

        hora = vals[0]

        # si está libre -> abrir formulario nueva cita
        if "libre" in tags:
            self._agenda_abrir_form_nueva_cita(hora_inicio=hora)
            return

        # si está ocupado -> abrir ventana de detalle
        if "ocupado" in tags:
            id_cita = tags[1]
            self._agenda_mostrar_detalle_cita(id_cita)

    def _agenda_mostrar_detalle_cita(self, id_cita):
        df = self.repo.dfs["Citas"]
        row = df[df["id_cita"] == id_cita]
        if row.empty:
            messagebox.showerror("Agenda", "No se encontró la cita.")
            return

        r = row.iloc[0]

        pid = r["id_paciente"]
        paciente = self._get_patient_name(pid)
        tel = self._get_patient_phone(pid)
        doctor = self._get_personal_name(r.get("id_doctor"))
        notas = r.get("notas", "")
        h1 = r.get("hora", "")
        hf = r.get("hora_fin", "")

        win = tk.Toplevel(self)
        win.title("Detalle de cita")
        win.geometry("360x280")
        win.grab_set()

        ttk.Label(win, text=f"Paciente: {paciente}", font=("Segoe UI",10,"bold")).pack(anchor="w", padx=10, pady=4)
        ttk.Label(win, text=f"Teléfono: {tel}").pack(anchor="w", padx=10)
        ttk.Label(win, text=f"Doctor: {doctor}").pack(anchor="w", padx=10)
        ttk.Label(win, text=f"Horario: {h1} - {hf}").pack(anchor="w", padx=10, pady=4)
        ttk.Label(win, text="Notas:").pack(anchor="w", padx=10, pady=(10,0))

        txt = tk.Text(win, height=6, width=40)
        txt.insert("1.0", notas)
        txt.config(state="disabled")
        txt.pack(padx=10, pady=4)

        ttk.Button(win, text="Cerrar", command=win.destroy).pack(pady=10)

    def _agenda_abrir_form_nueva_cita(self, paciente_id=None, hora_inicio=None):
        """Formulario avanzado para agendar cita con intervalos de 15 minutos."""
        win = tk.Toplevel(self)
        win.title("Nueva cita")
        win.geometry("420x460")
        win.grab_set()

        frm = ttk.Frame(win)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        # --------------------------------------------------------
        # VARIABLES
        # --------------------------------------------------------
        var_paciente = tk.StringVar()
        var_telefono = tk.StringVar()

        var_fecha = tk.StringVar(value=date.today().isoformat())
        var_doctor = tk.StringVar()
        var_notas = tk.StringVar()

        # Hora inicio por defecto = parámetro o 10:00
        var_hora_inicio = tk.StringVar(value=hora_inicio or "10:00")
        var_hora_fin = tk.StringVar(value="10:30")  # valor por defecto

        # Si viene desde finalizar consulta, autollenar
        if paciente_id:
            var_paciente.set(self._get_patient_name(paciente_id))
            var_telefono.set(self._get_patient_phone(paciente_id))
            var_pid = paciente_id
        else:
            var_pid = None

        # --------------------------------------------------------
        # GENERADOR DE HORAS (intervalos 15 min)
        # --------------------------------------------------------
        def generar_intervalos():
            horas = []
            h = 8
            m = 0
            while h < 20:  # hasta 19:45
                horas.append(f"{h:02d}:{m:02d}")
                m += 15
                if m == 60:
                    m = 0
                    h += 1
            return horas

        opciones_horas = generar_intervalos()

        # --------------------------------------------------------
        # CAMPOS
        # --------------------------------------------------------
        def row(lbl):
            f = ttk.Frame(frm)
            f.pack(fill="x", pady=4)
            ttk.Label(f, text=lbl, width=12, anchor="e").pack(side="left")
            return f

        # Paciente
        f = row("Paciente:")
        ttk.Entry(f, textvariable=var_paciente, width=28).pack(side="left", padx=6)

        f = row("Teléfono:")
        ttk.Entry(f, textvariable=var_telefono, width=18).pack(side="left", padx=6)

        f = row("Fecha:")
        ttk.Entry(f, textvariable=var_fecha, width=12).pack(side="left", padx=6)

        # Hora inicio
        """f = row("Hora inicio:")
        cb_inicio = ttk.Combobox(f, textvariable=var_hora_inicio, values=opciones_horas, width=10, state="readonly")
        cb_inicio.pack(side="left", padx=6)"""

        def refrescar_horas_libres(*args):
            fecha_sel = var_fecha.get().strip()
            libres = self._agenda_get_horas_libres(fecha_sel)
            cb_inicio["values"] = libres

            # si la hora precargada ya no es válida, asignar la primera libre
            if var_hora_inicio.get() not in libres:
                if libres:
                    var_hora_inicio.set(libres[0])
                else:
                    var_hora_inicio.set("")

        # Combobox con horas libres
        cb_inicio = ttk.Combobox(f, textvariable=var_hora_inicio, width=10, state="readonly")
        cb_inicio.pack(side="left", padx=6)

        # cada que cambia la fecha, recalcular horas
        var_fecha.trace_add("write", refrescar_horas_libres)

        # inicializar valores
        refrescar_horas_libres()

        # Hora fin
        """f = row("Hora fin:")
        cb_fin = ttk.Combobox(f, textvariable=var_hora_fin, values=opciones_horas, width=10, state="readonly")
        cb_fin.pack(side="left", padx=6)"""

        def refrescar_horas_fin(*args):
            hi = var_hora_inicio.get()
            if not hi:
                cb_fin["values"] = []
                return

            # hora inicio + 15 min hasta el final del día
            h, m = map(int, hi.split(":"))
            t1 = datetime(2025, 1, 1, h, m)

            horas_posibles = []
            t = t1 + timedelta(minutes=15)
            while t.hour < 20:
                horas_posibles.append(t.strftime("%H:%M"))
                t += timedelta(minutes=15)

            cb_fin["values"] = horas_posibles

            if var_hora_fin.get() not in horas_posibles:
                var_hora_fin.set(horas_posibles[0] if horas_posibles else "")

        cb_fin = ttk.Combobox(f, textvariable=var_hora_fin, width=10, state="readonly")
        cb_fin.pack(side="left", padx=6)

        var_hora_inicio.trace_add("write", refrescar_horas_fin)
        refrescar_horas_fin()

        f = row("Doctor:")
        ttk.Entry(f, textvariable=var_doctor, width=20).pack(side="left", padx=6)

        f = row("Notas:")
        ttk.Entry(f, textvariable=var_notas, width=28).pack(side="left", padx=6)

        # --------------------------------------------------------
        # VALIDACIONES
        # --------------------------------------------------------
        def validar_intervalo():
            """Valida que hora_fin sea mayor que hora_inicio."""
            hi = var_hora_inicio.get()
            hf = var_hora_fin.get()

            try:
                t1 = datetime.strptime(hi, "%H:%M")
                t2 = datetime.strptime(hf, "%H:%M")
            except:
                return False

            return t2 > t1

        def hay_conflicto_horario(pid, fecha, hi, hf):
            """Revisa si el horario se empalma con otras citas."""
            df = self.repo.citas_en_rango(fecha, fecha)
            if df.empty:
                return False

            ini_nueva = datetime.strptime(hi, "%H:%M").time()
            fin_nueva = datetime.strptime(hf, "%H:%M").time()

            for _, c in df.iterrows():
                ini = datetime.strptime(c["hora"], "%H:%M").time()

                # Duración de la cita existente (si tiene "hora_fin")
                fin = None
                if "hora_fin" in c and isinstance(c["hora_fin"], str) and c["hora_fin"]:
                    fin = datetime.strptime(c["hora_fin"], "%H:%M").time()

                if not fin:
                    # si no hay hora_fin, asumimos 30 min por defecto
                    t = datetime.combine(date.today(), ini) + timedelta(minutes=30)
                    fin = t.time()

                # Revisa empalme:
                if not (fin_nueva <= ini or ini_nueva >= fin):
                    return True

            return False

        # --------------------------------------------------------
        # GUARDAR CITA
        # --------------------------------------------------------
        def guardar():
            nombre = var_paciente.get().strip()
            tel = var_telefono.get().strip()
            fecha_txt = var_fecha.get().strip()
            hi = var_hora_inicio.get().strip()
            hf = var_hora_fin.get().strip()
            doc = var_doctor.get().strip()
            notas = var_notas.get().strip()

            if not nombre or not hi or not hf or not fecha_txt:
                messagebox.showwarning("Citas", "Faltan datos obligatorios.")
                return

            if not validar_intervalo():
                messagebox.showwarning("Citas", "La hora final debe ser mayor que la inicial.")
                return

            try:
                fecha_dt = datetime.strptime(fecha_txt, "%Y-%m-%d").date()
            except:
                messagebox.showwarning("Citas", "Fecha inválida.")
                return

            # Si no tiene paciente_id se crea o recupera
            pid = var_pid if var_pid else self.repo.upsert_patient(nombre, tel)[0]

            # Validar choques
            if hay_conflicto_horario(pid, fecha_dt, hi, hf):
                messagebox.showerror("Citas", "Hay un empalme de horarios.")
                return

            id_doc = self.repo.ensure_personal(doc, "doctor") if doc else ""

            # Crear cita con hora_fin incluida
            self.repo.crear_cita(pid, fecha_dt, hi, id_doctor=id_doc, notas=notas, hora_fin=hf)

            messagebox.showinfo("Citas", "Cita guardada.")
            win.destroy()
            try:
                self._load_citas_dia()
            except:
                pass

        ttk.Button(frm, text="Guardar cita", command=guardar).pack(pady=12)
     


    # -------------------------
    # Registro tab (con items)
    # -------------------------
    
    def _build_registro_tab(self):
        frm = self.tab_registro
        form = ttk.Frame(frm)
        form.pack(fill="x", padx=12, pady=8)

        # Variables
        self._var_paciente = tk.StringVar()
        self._var_telefono = tk.StringVar()
        self._var_doctor = tk.StringVar()
        self._var_promotor = tk.StringVar()

        
        # Campos
        def add_row(parent, label, var, width=40):
            row = ttk.Frame(parent)
            row.pack(fill="x", pady=4)

            ttk.Label(row, text=label, width=16, anchor="e").pack(side="left")

            ent = ttk.Entry(row, textvariable=var, width=width)
            ent.pack(side="left", padx=6)

            return ent

        self.ent_paciente = add_row(form, "Paciente", self._var_paciente)
        add_row(form, "Teléfono", self._var_telefono, width=28)

        self.lbl_cita = ttk.Label(form, text="", foreground="#B45F04")
        self.lbl_cita.pack(anchor="w", pady=2)

        self._var_paciente.trace_add("write", lambda *a: self._buscar_paciente_existente())
        self._var_telefono.trace_add("write", lambda *a: self._buscar_paciente_existente())

        add_row(form, "Doctor", self._var_doctor, width=28)
        add_row(form, "Promotor", self._var_promotor, width=28)

        self.ent_paciente.focus_set()
        # Botón para agregar a lista de espera
        btn_frame = ttk.Frame(form)
        btn_frame.pack(fill="x", pady=10)
        ttk.Button(btn_frame, text="Agregar a lista de espera", command=self._guardar_en_espera).pack(side="left", padx=6)

        # Etiqueta de salida
        self.lbl_reg_out = ttk.Label(frm, text="")
        self.lbl_reg_out.pack(anchor="w", padx=12, pady=6)

    def _buscar_paciente_existente(self):
        nombre = self._var_paciente.get().strip()
        telefono = self._var_telefono.get().strip()

        df = self.repo.dfs["Pacientes"]

        row = None

        if telefono:
            match = df[df["telefono"] == telefono]
            if not match.empty:
                row = match.iloc[0]

        if row is None and nombre:
            match = df[df["nombre"].str.lower() == nombre.lower()]
            if not match.empty:
                row = match.iloc[0]

        if row is None:
            self.lbl_folio.config(text="")
            self.paciente_id_actual = None
            return

        self.paciente_id_actual = row["id_paciente"]

        folio = row.get("folio_paciente", "")
        if folio:
            self.lbl_folio.config(text=f"Expediente: {folio}")
        else:
            self.lbl_folio.config(text="(Paciente sin expediente)")

##################################  CAMBIOS EN ESPERA  ##################################
  
    def _guardar_en_espera(self):
        if not self._ensure_repo():
            return

        paciente = self._var_paciente.get().strip()
        telefono = self._var_telefono.get().strip()
        doctor = self._var_doctor.get().strip() or None
        promotor = self._var_promotor.get().strip() or None

        if not paciente:
            messagebox.showwarning("Registro", "El nombre del paciente es obligatorio.")
            return

        try:
            # Crear o actualizar paciente
            pid, _, _ = self.repo.upsert_patient(paciente, telefono)

            # Registrar personal (doctor y promotor)
            id_doc = self.repo.ensure_personal(doctor, "doctor") if doctor else ""
            id_prom = self.repo.ensure_personal(promotor, "promotor") if promotor else ""

            # Crear fila en Consultas
            df = self.repo.dfs["Consultas"]
            nuevo_id = f"C{datetime.now().strftime('%y%m%d%H%M%S')}"
            nueva = {
                "id_consulta": nuevo_id,
                "id_paciente": pid,
                "fecha": datetime.now(),
                "id_doctor": id_doc,
                "id_promotor": id_prom,
                "estado": "en espera",
                "total_consulta": 0.0
            }
            if "metodo_pago" in df.columns:
                nueva["metodo_pago"] = ""
            if "moneda" in df.columns:
                nueva["moneda"] = ""
            if "notas" in df.columns:
                nueva["notas"] = ""

            self.repo.dfs["Consultas"] = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
            self.repo.save()

            messagebox.showinfo("Registro", f"Consulta en espera creada correctamente.")
            self._var_paciente.set("")
            self._var_telefono.set("")
            self._var_doctor.set("")
            self._var_promotor.set("")
            try:
                self._load_consultas_espera()
            except:
                pass
        except Exception as e:
            messagebox.showerror("Registro", f"Error al guardar consulta en espera:\n{e}")

    def _build_espera_tab(self):
        frm = self.tab_espera

        top = ttk.Frame(frm)
        top.pack(fill="x", padx=12, pady=8)

        ttk.Label(top, text="Consultas en espera / consultando").pack(anchor="w")

        # ======== 1️⃣ Declaramos columnas (incluyendo id_consulta) ========
        cols = ("id_consulta", "folio", "fecha", "paciente", "telefono", "doctor", "promotor", "estado")

        # ======== 2️⃣ Crear el TreeView PRIMERO ========
        tree = ttk.Treeview(frm, columns=cols, show="headings", height=12)

        # ======== 3️⃣ Configurar encabezados y columnas ========
        for c in cols:
            tree.heading(c, text=c)
            if c == "id_consulta":
                # Columna oculta
                tree.column(c, width=0, stretch=False)
            elif c == "paciente":
                tree.column(c, width=180, anchor="center")
            else:
                tree.column(c, width=130, anchor="center")

        # ======== 4️⃣ Envolver con scroll ========
        self.tree_espera = self._wrap_tree_with_scroll(frm, tree)

        # ======== 5️⃣ Zebra stripes ========
        self._apply_zebra(self.tree_espera)

        # ======== Botones ========
        btns = ttk.Frame(frm)
        btns.pack(fill="x", padx=12, pady=8)

        ttk.Button(btns, text="Refrescar lista", command=self._load_consultas_espera).pack(side="left")
        ttk.Button(btns, text="Iniciar consulta", command=lambda: self._change_estado_consulta("consultando")).pack(side="left", padx=8)
        ttk.Button(btns, text="Cancelar", command=lambda: self._change_estado_consulta("cancelada")).pack(side="left", padx=8)
        ttk.Button(btns, text="Finalizar consulta", command=self._abrir_finalizar_consulta).pack(side="right", padx=8)

    def _load_consultas_espera(self):
        if not self._ensure_repo(): 
            return

        df = self.repo.dfs["Consultas"]
        if df.empty: 
            return

        df = df[df["estado"].isin(["en espera", "consultando"])]

        rows = []
        tags = []

        for _, r in df.iterrows():
            pid = r.get("id_paciente")
            vals = (
                r["id_consulta"],             # <-- NUEVO
                self._get_patient_folio(pid),
                str(r.get("fecha",""))[:19],
                self._get_patient_name(pid),
                self._get_patient_phone(pid),
                self._get_personal_name(r.get("id_doctor")),
                self._get_personal_name(r.get("id_promotor")),
                r.get("estado","")
            )
            rows.append(vals)
            tags.append((r["id_consulta"],))

        self.tree_espera.delete(*self.tree_espera.get_children())

        for i, vals in enumerate(rows):
            tag = "oddrow" if i % 2 else "evenrow"
            real_id = tags[i]
            self.tree_espera.insert("", "end", values=vals, tags=(tag,))

    def _get_patient_name(self, pid):
        df = self.repo.dfs["Pacientes"]
        row = df[df["id_paciente"] == pid]
        return row.iloc[0]["nombre"] if not row.empty else ""

    def _get_patient_phone(self, pid):
        df = self.repo.dfs["Pacientes"]
        row = df[df["id_paciente"] == pid]
        return row.iloc[0]["telefono"] if not row.empty else ""
    
    def _get_patient_folio(self, pid):
        """Devuelve el folio del paciente o cadena vacía."""
        try:
            df = self.repo.dfs["Pacientes"]
            row = df[df["id_paciente"] == pid]
            if not row.empty:
                folio = row.iloc[0].get("folio_paciente", "")
                return folio if isinstance(folio, str) else ""
        except Exception:
            pass
        return ""


    def _get_personal_name(self, pid):
        if not pid: return ""
        df = self.repo.dfs["Personal"]
        row = df[df["id_personal"] == pid]
        return row.iloc[0]["nombre"] if not row.empty else ""

    def _change_estado_consulta(self, nuevo_estado):
        sel = self.tree_espera.selection()
        if not sel:
            messagebox.showwarning("Consultas", "Selecciona una consulta.")
            return

        iid = sel[0]

        # Obtener el ID REAL desde tags (NO desde values)
        vals = self.tree_espera.item(iid)["values"]
        id_consulta = vals[0]   # ahora SIEMPRE existe

        df = self.repo.dfs["Consultas"]
        idx = df.index[df["id_consulta"] == id_consulta].tolist()
        if not idx:
            messagebox.showerror("Consultas", "No se encontró la consulta seleccionada.")
            return

        self.repo.dfs["Consultas"].at[idx[0], "estado"] = nuevo_estado
        self.repo.save()
        self._load_consultas_espera()
        messagebox.showinfo("Consultas", f"El estado se cambió a '{nuevo_estado}'.")

    def _abrir_finalizar_consulta(self):
        sel = self.tree_espera.selection()
        if not sel:
            messagebox.showwarning("Consultas", "Selecciona una consulta para finalizar.")
            return

        item_id = sel[0]

        
        vals = self.tree_espera.item(item_id)["values"]
        id_consulta = vals[0]
        # Ahora sí obtenemos el estado real desde la tabla Consultas
        df = self.repo.dfs["Consultas"]
        row = df[df["id_consulta"] == id_consulta]
        if row.empty:
            messagebox.showerror("Error", "No se encontró la consulta seleccionada.")
            return

        estado_actual = row.iloc[0]["estado"].strip().lower()

        if estado_actual == "en espera":
            messagebox.showwarning("Consulta no iniciada", "Primero cambia su estado a 'consultando'.")
            return

        if estado_actual != "consultando":
            messagebox.showwarning("Estado inválido", "Solo las consultas en estado 'consultando' pueden finalizarse.")
            return

        self._abrir_form_finalizacion(id_consulta)

    def _abrir_form_finalizacion(self, id_consulta):
        win = tk.Toplevel(self)
        win.title(f"Finalizar consulta {id_consulta}")
        win.geometry("520x540")
        win.grab_set()

        ttk.Label(win, text=f"Consulta: {id_consulta}", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=5)

        frm = ttk.Frame(win)
        frm.pack(fill="x", padx=12, pady=8)

        # Variables
        var_tecnico = tk.StringVar()
        var_costo_tecnico = tk.StringVar()
        var_tratamiento = tk.StringVar()
        var_cantidad = tk.StringVar(value="1")
        var_precio = tk.StringVar(value="0.0")
        # --- Método de pago (Combobox) ---
        var_metodo = tk.StringVar()
        ttk.Label(frm, text="Método de pago:", anchor="e", width=16).pack(anchor="w")
        cmb_metodo = ttk.Combobox(
            frm, textvariable=var_metodo, state="readonly",
            values=["Efectivo", "Tarjeta", "Transferencia", "Zelle", "Otro"]
        )
        cmb_metodo.current(0)
        cmb_metodo.pack(anchor="w", padx=6, pady=3)

        # --- Moneda (radiobuttons estilo checklist) ---
        ttk.Label(frm, text="Moneda:", anchor="e", width=16).pack(anchor="w", pady=(6, 0))
        var_moneda = tk.StringVar(value="USD")
        frm_mon = ttk.Frame(frm)
        frm_mon.pack(anchor="w", padx=12)
        for m in ["USD", "MXN"]:
            ttk.Radiobutton(frm_mon, text=m, variable=var_moneda, value=m).pack(side="left", padx=8)

        # --- Folio clínico (ligado al paciente) ---
        var_folio = tk.StringVar()
        folio_readonly = False

        try:
            # Cargar folio actual del paciente (si ya existe)
            df_cons = self.repo.dfs["Consultas"]
            row = df_cons[df_cons["id_consulta"] == id_consulta].iloc[0]
            id_paciente = row["id_paciente"]
            pac = self.repo.dfs["Pacientes"]
            folio_exist = pac.loc[pac["id_paciente"] == id_paciente, "folio_paciente"].iloc[0]
            if isinstance(folio_exist, str) and folio_exist.strip() != "":
                var_folio.set(folio_exist)
                folio_readonly = True
        except Exception:
            pass

        # Mostrar campo de folio
        f_folio = ttk.Frame(frm)
        f_folio.pack(fill="x", pady=4)
        ttk.Label(f_folio, text="Folio clínico:", width=16, anchor="e").pack(side="left")

        ent_folio = ttk.Entry(f_folio, textvariable=var_folio, width=30)
        ent_folio.pack(side="left", padx=6)

        # Si ya tiene folio, dejarlo solo lectura
        if folio_readonly:
            ent_folio.state(["readonly"])
        var_notas = tk.StringVar()
        var_agendar = tk.BooleanVar(value=False)

        # Campos principales
        def add_row(lbl, var, width=30):
            f = ttk.Frame(frm)
            f.pack(fill="x", pady=4)
            ttk.Label(f, text=lbl, width=16, anchor="e").pack(side="left")
            ttk.Entry(f, textvariable=var, width=width).pack(side="left", padx=6)

        add_row("Técnico:", var_tecnico)
        add_row("Costo técnico:", var_costo_tecnico)
        ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=6)
        add_row("Tratamiento:", var_tratamiento)
        add_row("Cantidad:", var_cantidad, width=10)
        add_row("Precio unitario:", var_precio, width=10)
        ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=6)
        
        ttk.Label(frm, text="Notas:").pack(anchor="w")
        txt_notas = tk.Text(frm, height=4, width=48)
        txt_notas.pack(pady=4)

        ttk.Checkbutton(frm, text="Agendar próxima cita", variable=var_agendar).pack(anchor="w", pady=6)

        def guardar_finalizacion():
            if not self._ensure_repo():
                return
            try:
                tecnico = var_tecnico.get().strip()
                costo_tecnico = float(var_costo_tecnico.get() or 0)
                tratamiento = var_tratamiento.get().strip()
                cantidad = float(var_cantidad.get() or 1)
                precio = float(var_precio.get() or 0)
                metodo = var_metodo.get().strip()
                moneda = var_moneda.get().strip().upper()
                notas = txt_notas.get("1.0", "end").strip()
                folio = var_folio.get().strip()

                if not tratamiento:
                    messagebox.showwarning("Finalizar consulta", "Debe ingresar un tratamiento.")
                    return

                subtotal = cantidad * precio

                # Guardar item
                df_items = self.repo.dfs["Consulta_Items"]
                iid = f"I{datetime.now().strftime('%y%m%d%H%M%S')}"
                nuevo_item = {
                    "id_item": iid,
                    "id_consulta": id_consulta,
                    "id_tratamiento": "",  # se resolverá en backend
                    "descripcion_item": tratamiento,
                    "cantidad": cantidad,
                    "precio_unitario": precio,
                    "subtotal": subtotal,
                    "id_tecnico_item": self.repo.ensure_personal(tecnico, "tecnico") if tecnico else "",
                }
                self.repo.dfs["Consulta_Items"] = pd.concat([df_items, pd.DataFrame([nuevo_item])], ignore_index=True)

                # Actualizar cabecera de consulta
                df_cons = self.repo.dfs["Consultas"]
                idx = df_cons.index[df_cons["id_consulta"] == id_consulta].tolist()
                if not idx:
                    messagebox.showerror("Error", "No se encontró la consulta seleccionada.")
                    return

                row = df_cons[df_cons["id_consulta"] == id_consulta].iloc[0]
                id_paciente = row["id_paciente"]

                # Validar si no hay folio y mostrar advertencia
                if not folio:
                    resp = messagebox.askyesno(
                        "Folio clínico faltante",
                        "No has ingresado el número de folio clínico del paciente.\n"
                        "¿Deseas continuar y guardar SIN folio clínico?\n\n"
                        "Selecciona 'Sí' para continuar sin folio o 'No' para regresar al formulario."
                    )
                    if not resp:
                        # Cancelar guardado
                        return
                else:
                    # Guardar folio en el paciente si es nuevo
                    self.repo.update_paciente_folio(id_paciente, folio)

                df_cons.at[idx[0], "total_consulta"] = subtotal + costo_tecnico
                df_cons.at[idx[0], "metodo_pago"] = metodo
                df_cons.at[idx[0], "moneda"] = moneda
                df_cons.at[idx[0], "notas"] = notas
                df_cons.at[idx[0], "estado"] = "finalizada"

                self.repo.save()
                self._load_consultas_espera()
                messagebox.showinfo("Consulta finalizada", "Los datos se guardaron correctamente.")

                if var_agendar.get():

                    self._agenda_abrir_form_nueva_cita(paciente_id=id_paciente)

                win.destroy()
            except Exception as e:
                messagebox.showerror("Finalizar consulta", f"Error al guardar:\n{e}")

        ttk.Button(frm, text="Guardar finalización", command=guardar_finalizacion).pack(pady=10)

#########################################################################################    
    def _agregar_item(self):
        desc = self._item_tratamiento.get().strip()
        try:
            cant = float(self._item_cantidad.get())
            precio = float(self._item_precio.get())
        except:
            messagebox.showwarning("Item", "Cantidad y precio deben ser numéricos.")
            return
        if not desc or cant<=0:
            messagebox.showwarning("Item", "Describe el tratamiento y cantidad > 0.")
            return
        subtotal = round(cant*precio,2)
        # id_tratamiento vacío por ahora (se resolverá en backend)
        self.tree_items.insert("", "end", values=(desc, f"{cant:.2f}", f"{precio:.2f}", f"{subtotal:.2f}", ""))
        # limpiar inputs
        self._item_tratamiento.set(""); self._item_cantidad.set("1"); self._item_precio.set("0.0"); self._item_desc.set("")

    def _eliminar_item(self):
        sel = self.tree_items.selection()
        if not sel: return
        for s in sel: self.tree_items.delete(s)



    def _ensure_repo(self):
        if not self.cfg:
            messagebox.showwarning("Config","Carga config.json en Config tab.")
            return False
        if not self.repo:
            if not self.cfg.excel_path.exists():
                messagebox.showwarning("Excel","Crea plantilla primero en Config.")
                return False
            self.repo = Repo(self.cfg)
            self._comision_service = ComisionService(self.repo, self.cfg)
        elif not self._comision_service:
            self._comision_service = ComisionService(self.repo, self.cfg)
        return True

    # -------------------------
    # Historial (simple)
    # -------------------------
    def _build_historial_tab(self):
        frm = self.tab_historial
        top = ttk.Frame(frm); top.pack(fill="x", padx=12, pady=8)
        self._hist_by = tk.StringVar(value="nombre"); self._hist_q = tk.StringVar()
        ttk.Label(top, text="Buscar por:").pack(side="left"); ttk.Radiobutton(top, text="Nombre", value="nombre", variable=self._hist_by).pack(side="left", padx=4); ttk.Radiobutton(top, text="Teléfono", value="telefono", variable=self._hist_by).pack(side="left", padx=4)
        ttk.Entry(top, textvariable=self._hist_q, width=40).pack(side="left", padx=8); ttk.Button(top, text="Buscar", command=self._do_historial).pack(side="left")
        cols = ("id_consulta","fecha","total_consulta","estado")
        self.tree_hist = ttk.Treeview(frm, columns=cols, show="headings", height=20)
        for c in cols: self.tree_hist.heading(c,text=c); self.tree_hist.column(c,width=140,anchor="center")
        self.tree_hist.pack(fill="both", expand=True, padx=12, pady=8)
        self.lbl_hist_info = ttk.Label(frm,text=""); self.lbl_hist_info.pack(anchor="w", padx=12, pady=4)

    def _do_historial(self):
        if not self._ensure_repo(): return
        q = self._hist_q.get().strip(); by = self._hist_by.get()
        if not q:
            messagebox.showwarning("Historial","Ingresa texto")
            return
        try:
            if by=="nombre":
                hits = self.repo.find_patient_by_name(q)
            else:
                hits = self.repo.find_patient_by_phone(q)
            if hits.empty:
                messagebox.showinfo("Historial","Sin coincidencias")
                self.tree_hist.delete(*self.tree_hist.get_children())
                self.lbl_hist_info.config(text="")
                return

            pid = hits.iloc[0]["id_paciente"]
            cons = self.repo.dfs["Consultas"]
            cons = cons[cons["id_paciente"] == pid].sort_values("fecha", ascending=False)
            self.tree_hist.delete(*self.tree_hist.get_children())

            for _, r in cons.iterrows():
                self.tree_hist.insert("", "end", values=(
                    r.get("id_consulta", ""),
                    str(r.get("fecha", ""))[:19],
                    r.get("total_consulta", ""),
                    r.get("estado", "")
                ))

            self.lbl_hist_info.config(
                text=f"Paciente: {hits.iloc[0]['nombre']}  | Tel: {hits.iloc[0].get('telefono','')}"
            )
        except Exception as e:
            messagebox.showerror("Historial", f"Error:\n{e}")

    # -------------------------
    # Reportes (usa items)
    # -------------------------
    def _build_reportes_tab(self):
        frm = self.tab_reportes

        # ---- Filtros superiores ----
        filters = ttk.Frame(frm)
        filters.pack(fill="x", padx=12, pady=8)

        self._rep_desde = tk.StringVar()
        self._rep_hasta = tk.StringVar()
        self._rep_paciente = tk.StringVar()
        self._rep_doctor = tk.StringVar()
        self._rep_promotor = tk.StringVar()
        self._rep_tecnico = tk.StringVar()
        self._rep_tratamiento = tk.StringVar()

        # Fecha desde
        box_desde = ttk.Frame(filters)
        box_desde.pack(side="left", padx=6)
        ttk.Label(box_desde, text="Desde").pack(anchor="w")
        de_desde = self._mk_date_selector(box_desde, var=self._rep_desde)
        de_desde.pack(anchor="w")

        # Fecha hasta
        box_hasta = ttk.Frame(filters)
        box_hasta.pack(side="left", padx=6)
        ttk.Label(box_hasta, text="Hasta").pack(anchor="w")
        de_hasta = self._mk_date_selector(box_hasta, var=self._rep_hasta)
        de_hasta.pack(anchor="w")

        # Filtros de texto
        def add_filter(label, var, width=18):
            box = ttk.Frame(filters)
            box.pack(side="left", padx=6)
            ttk.Label(box, text=label).pack(anchor="w")
            ttk.Entry(box, textvariable=var, width=width).pack(anchor="w")

        add_filter("Paciente",   self._rep_paciente)
        add_filter("Doctor",     self._rep_doctor)
        add_filter("Promotor",   self._rep_promotor)
        add_filter("Técnico",    self._rep_tecnico)
        add_filter("Tratamiento",self._rep_tratamiento)

        ttk.Button(filters, text="Filtrar", command=self._do_reportes).pack(side="left", padx=10)

        # ---- Contenedor para la tabla ----
        cont_rep = ttk.Frame(frm)
        cont_rep.pack(fill="both", expand=True, padx=12, pady=8)

        self._rep_cols = (
            "fecha",
            "paciente",
            "tratamiento",
            "cantidad",
            "precio_unitario",
            "subtotal",
            "doc_nombre",
            "pro_nombre",
            "tec_nombre",
        )

        # Treeview base (se recreará internamente en _wrap_tree_with_scroll)
        self.tree_rep = ttk.Treeview(cont_rep, columns=self._rep_cols, show="headings", height=20)

        # Envolver con scroll y layout correcto
        self.tree_rep = self._wrap_tree_with_scroll(cont_rep, self.tree_rep)

        # Configurar columnas/encabezados
        for c in self._rep_cols:
            self.tree_rep.heading(c, text=c)
            width = 120
            if c in ("paciente", "tratamiento"):
                width = 160
            self.tree_rep.column(c, width=width, anchor="center")

        self._apply_zebra(self.tree_rep)
        # Evento de clic para mostrar notas del paciente
        self.tree_rep.bind("<Double-1>", self._mostrar_notas_paciente)

        

    def _do_reportes(self):
        if not self._ensure_repo(): return
        try:
            desde = _parse_date(self._rep_desde.get()) if self._rep_desde.get().strip() else None
            hasta = _parse_date(self._rep_hasta.get()) if self._rep_hasta.get().strip() else None
            rep = ReporteService(self.repo, self.cfg)
            df = rep.filtrar(desde, hasta, self._rep_paciente.get().strip() or None, self._rep_doctor.get().strip() or None, self._rep_promotor.get().strip() or None, self._rep_tecnico.get().strip() or None, self._rep_tratamiento.get().strip() or None)
            
            if df.empty:
                messagebox.showinfo("Reportes","Sin resultados")
                return

            df = df.copy()
            df["tratamiento"] = df["tratamiento"].fillna(df["descripcion_item"])

            rows = []
            for _, r in df.sort_values("fecha").iterrows():
                vals = []
                for c in self._rep_cols:
                    raw = r.get(c, "")
                    sval = "" if str(raw).lower() in ("nan","none") else str(raw)
                    if c == "tratamiento":
                        sval = sval.split(" - ",1)[-1]
                    vals.append(sval)
                rows.append(vals)

            self._zebra_insert_all(self.tree_rep, rows)
        except Exception as e:
            messagebox.showerror("Reportes", f"Error: {e}")



 # -------------------------
    # Cortes (selección y detalle)
    # -------------------------
    def _build_cortes_tab(self):
        frm = self.tab_cortes
        top = ttk.Frame(frm); top.pack(fill="x", padx=12, pady=8)
        self._corte_fecha = tk.StringVar(); self._corte_pdf = tk.BooleanVar(value=True)
        #ttk.Label(top, text="Fecha (YYYY-MM-DD):").pack(side="left"); de_corte = self._mk_date_entry(top, var=self._corte_fecha, default=date.today()); de_corte.pack(side="left", padx=6)
        
        ttk.Label(top, text="Fecha (YYYY-MM-DD):").pack(side="left")
        de_corte = self._mk_date_selector(top, var=self._corte_fecha, default=date.today())
        de_corte.pack(side="left", padx=6)

        ttk.Checkbutton(top, text="Generar PDF", variable=self._corte_pdf).pack(side="left", padx=10); ttk.Button(top, text="Generar corte", command=self._do_corte).pack(side="left", padx=10)
        ttk.Separator(frm,orient="horizontal").pack(fill="x", padx=12, pady=6)
        range_box = ttk.Frame(frm); range_box.pack(fill="x", padx=12, pady=6)
        self._cortes_desde = tk.StringVar(); self._cortes_hasta = tk.StringVar()
        #ttk.Label(range_box, text="Desde").pack(side="left"); de_c_desde = self._mk_date_entry(range_box, var=self._cortes_desde); de_c_desde.pack(side="left", padx=4)
        #ttk.Label(range_box, text="Hasta").pack(side="left"); de_c_hasta = self._mk_date_entry(range_box, var=self._cortes_hasta); de_c_hasta.pack(side="left", padx=4)
        
        de_c_desde = self._mk_date_selector(range_box, var=self._cortes_desde)
        de_c_desde.pack(side="left", padx=4)

        de_c_hasta = self._mk_date_selector(range_box, var=self._cortes_hasta)
        de_c_hasta.pack(side="left", padx=4)
        
        ttk.Button(range_box, text="Consultar", command=self._do_cortes_lista).pack(side="left", padx=8)
        cols = ("id_corte","fecha_corte","rango_inicio","rango_fin","total_ingresos","total_utilidad","creado_en")
        self.tree_cortes = ttk.Treeview(frm, columns=cols, show="headings", height=6)
        for c in cols: self.tree_cortes.heading(c,text=c); self.tree_cortes.column(c,width=150 if c in ("rango_inicio","rango_fin") else 120, anchor="center")
        
        self.tree_cortes = self._wrap_tree_with_scroll(frm, self.tree_cortes)
        self._apply_zebra(self.tree_cortes)
        
        self.tree_cortes.bind("<<TreeviewSelect>>", self._on_corte_select)
        bottom = ttk.Frame(frm); bottom.pack(fill="both", expand=True, padx=12, pady=6)
        left = ttk.LabelFrame(bottom, text="Resumen por persona"); left.pack(side="left", fill="y", padx=(0,8))
        
        
        # === DOCTORES ===
        ttk.Label(left, text="Doctores (a pagar)").pack(anchor="w")
        self.tree_res_doc = ttk.Treeview(left, columns=("nombre","total"), show="headings", height=6)

        self.tree_res_doc.heading("nombre", text="nombre")
        self.tree_res_doc.heading("total", text="total")

        # Anchos compactos
        self.tree_res_doc.column("nombre", width=110, anchor="w")
        self.tree_res_doc.column("total", width=80, anchor="center")

        self.tree_res_doc = self._wrap_tree_with_scroll(left, self.tree_res_doc)
        self._apply_zebra(self.tree_res_doc)

        # === PROMOTORES ===
        ttk.Label(left, text="Promotores (a pagar)").pack(anchor="w")
        self.tree_res_pro = ttk.Treeview(left, columns=("nombre","total"), show="headings", height=6)

        self.tree_res_pro.heading("nombre", text="nombre")
        self.tree_res_pro.heading("total", text="total")

        self.tree_res_pro.column("nombre", width=110, anchor="w")
        self.tree_res_pro.column("total", width=80, anchor="center")

        self.tree_res_pro = self._wrap_tree_with_scroll(left, self.tree_res_pro)
        self._apply_zebra(self.tree_res_pro)

        # === TÉCNICOS ===
        ttk.Label(left, text="Técnicos (costo técnico)").pack(anchor="w")
        self.tree_res_tec = ttk.Treeview(left, columns=("nombre","total"), show="headings", height=6)

        self.tree_res_tec.heading("nombre", text="nombre")
        self.tree_res_tec.heading("total", text="total")

        self.tree_res_tec.column("nombre", width=110, anchor="w")
        self.tree_res_tec.column("total", width=80, anchor="center")

        self.tree_res_tec = self._wrap_tree_with_scroll(left, self.tree_res_tec)
        self._apply_zebra(self.tree_res_tec)
        
####################AQUI CREO LA TABLA DE REPORTES##########
        right = ttk.LabelFrame(bottom, text="Detalle del corte")
        right.pack(side="left", fill="both", expand=True)

        # columnas CORRECTAS del detalle
        cols_d = (
            "fecha","paciente","tratamiento",
            "doc_nombre","pro_nombre","tec_nombre",
            "subtotal","precio_unitario","cantidad",
            "comision_doctor","comision_promotor","utilidad"
        )

        # Crear tabla dentro de "right"
        self.tree_detalle = ttk.Treeview(right, columns=cols_d, show="headings", height=50)

        # Títulos de columnas
        for c in cols_d:
            self.tree_detalle.heading(c, text=c)

        # Configuración de ancho y alineación
        for c in cols_d:
            if c == "fecha":
                self.tree_detalle.column(c, width=115, anchor="center")
            elif c == "paciente":
                self.tree_detalle.column(c, width=130, anchor="w")
            elif c == "tratamiento":
                self.tree_detalle.column(c, width=140, anchor="w")  # era 170
            elif c in ("doc_nombre", "pro_nombre", "tec_nombre"):
                self.tree_detalle.column(c, width=110, anchor="center")
            elif c in ("subtotal","precio_unitario","cantidad",
                    "comision_doctor","comision_promotor","utilidad"):
                self.tree_detalle.column(c, width=90, anchor="center")

        self.tree_detalle = self._wrap_tree_with_scroll(right, self.tree_detalle)
        self._apply_zebra(self.tree_detalle)

    """def _do_corte(self):
        if not self._ensure_repo(): return
        try:
            fecha = _parse_date(self._corte_fecha.get().strip()) if self._corte_fecha.get().strip() else date.today()
            corte = CorteService(self.repo, self.cfg); id_corte, detalle, resumen = corte.generar_corte(fecha)
            msg = f"Corte generado: {id_corte}\n"
            for k,v in resumen.items(): msg += f"  {k}: {v}\n"
            if self._corte_pdf.get():
                from consultorio_items import export_pdf_corte
                try:
                    path = export_pdf_corte(self.cfg, id_corte, resumen, detalle)
                    msg += f"\nPDF: {path}"
                except Exception:
                    pass
            messagebox.showinfo("Corte", msg)
        except Exception as e:
            messagebox.showerror("Corte", f"Error al generar corte:\n{e}")"""


    def _do_corte(self):
        if not self._ensure_repo(): return
        try:
            fecha = _parse_date(self._corte_fecha.get().strip()) if self._corte_fecha.get().strip() else date.today()
            corte = CorteService(self.repo, self.cfg)
            id_corte, detalle, resumen = corte.generar_corte(fecha)
            msg = f"Corte generado: {id_corte}\n"
            for k,v in resumen.items(): msg += f"  {k}: {v}\n"
            if self._corte_pdf.get():
                from consultorio_items import export_pdf_corte
                try:
                    path = export_pdf_corte(self.cfg, id_corte, resumen, detalle)
                    msg += f"\nPDF: {path}"
                except Exception:
                    pass
            messagebox.showinfo("Corte", msg)
        except Exception as e:
            # Diagnóstico extendido: traceback + tipos de los objetos principales
            import traceback, pprint
            tb = traceback.format_exc()
            info = f"Error al generar corte:\n{e}\n\nTraceback:\n{tb}\n\n"
            try:
                # mostrar tipos y muestras de los dataframes en repo.dfs
                dfs = getattr(self.repo, "dfs", None)
                if isinstance(dfs, dict):
                    info += "Contenido de repo.dfs (tipo / preview):\n"
                    for k, v in dfs.items():
                        t = type(v)
                        info += f" - {k}: {t}\n"
                        try:
                            # si tiene head() es un DataFrame/Series
                            if hasattr(v, "head"):
                                head = v.head(3).to_dict(orient="records")
                                info += f"   head: {pprint.pformat(head)}\n"
                            else:
                                # si es string, mostrar longitud / muestra
                                if isinstance(v, str):
                                    info += f"   str(len={len(v)}): {v[:200]!r}\n"
                                else:
                                    info += f"   repr: {repr(v)[:200]}\n"
                        except Exception as inner:
                            info += f"   error al inspeccionar: {inner}\n"
                else:
                    info += f"repo.dfs no es un dict: {type(dfs)}\n"
            except Exception as diag_e:
                info += f"Error en diagnóstico: {diag_e}\n"

            # Mostrar en un cuadro grande para copiar/pegar
            messagebox.showerror("Corte - Error y diagnóstico", info)
            # además imprimir en consola (útil si ejecutas desde terminal)
            print(info)


    def _do_cortes_lista(self):
        if not self._ensure_repo(): return
        try:
            df = self.repo.dfs["Cortes"].copy()
            d = self._cortes_desde.get().strip(); h = self._cortes_hasta.get().strip()
            if d: df = df[df["fecha_corte"] >= d]
            if h: df = df[df["fecha_corte"] <= h]
            
            if df.empty:
                messagebox.showinfo("Cortes","No hay cortes")
                return

            rows = []
            for _, r in df.sort_values("fecha_corte").iterrows():
                rows.append((
                    r.get("id_corte",""),
                    str(r.get("fecha_corte","")),
                    r.get("rango_inicio",""),
                    r.get("rango_fin",""),
                    r.get("total_ingresos",""),
                    r.get("total_utilidad",""),
                    r.get("creado_en","")
                ))

            self._zebra_insert_all(self.tree_cortes, rows)

        except Exception as e:
            messagebox.showerror("Cortes", f"Error:\n{e}")

    def _mostrar_notas_paciente(self, event):
        """Abre una ventana emergente mostrando el historial de notas del paciente."""
        if not self._ensure_repo():
            return

        # Selección
        sel = self.tree_rep.selection()
        if not sel:
            return

        item = sel[0]
        valores = self.tree_rep.item(item)["values"]

        # columnas en tree_rep:
        # ("fecha","paciente","tratamiento","cantidad","precio_unitario",
        #  "subtotal", "doc_nombre","pro_nombre","tec_nombre")
        paciente_nombre = valores[1]

        # Obtener ID del paciente
        df_pac = self.repo.dfs["Pacientes"]
        row = df_pac[df_pac["nombre"].str.lower() == paciente_nombre.lower()]
        if row.empty:
            messagebox.showwarning("Notas", "No se encontró el paciente.")
            return

        id_paciente = row.iloc[0]["id_paciente"]

        # Obtener historial de notas desde backend
        notas_df = self.repo.get_notas_paciente(id_paciente)

        # Crear ventana emergente
        win = tk.Toplevel(self)
        win.title(f"Notas de {paciente_nombre}")
        win.geometry("600x400")
        win.grab_set()

        ttk.Label(win, text=f"Historial de notas de: {paciente_nombre}",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=10, pady=8)

        # Textbox para mostrar notas
        txt = tk.Text(win, height=20, width=80)
        txt.pack(fill="both", expand=True, padx=10, pady=10)

        if notas_df.empty:
            txt.insert("end", "El paciente no tiene notas guardadas.")
        else:
            for _, r in notas_df.iterrows():
                fecha = str(r["fecha"])[:19]
                notas = r["notas"]
                txt.insert("end", f"📌 {fecha}\n{notas}\n\n")

        txt.config(state="disabled")

        ttk.Button(win, text="Cerrar", command=win.destroy).pack(pady=10)



    def _on_corte_select(self, event):
        try:
            sel = self.tree_cortes.selection()
            if not sel:
                return

            item = self.tree_cortes.item(sel[0])
            id_corte = item["values"][0]

            corte_service = CorteService(self.repo, self.cfg)
            df, resumen = corte_service.detalle_corte(id_corte)

            # --- LIMPIAR TABLAS ---
            self.tree_res_doc.delete(*self.tree_res_doc.get_children())
            self.tree_res_pro.delete(*self.tree_res_pro.get_children())
            self.tree_res_tec.delete(*self.tree_res_tec.get_children())
            self.tree_detalle.delete(*self.tree_detalle.get_children())

            # --- LIMPIAR NANs EN DETALLE COMPLETO ---
            df = df.replace({pd.NA: "", None: "", "nan": "", "NaN": ""})

            # =============================================================
            # 1) RESUMEN: DOCTORES
            # =============================================================
            rows_doc = []
            doctores = resumen.get("doctores", {})

            if isinstance(doctores, dict):
                for nombre, total in doctores.items():
                    rows_doc.append([
                        "" if str(nombre).lower() == "nan" else nombre,
                        total
                    ])
            elif hasattr(doctores, "items"):   # pandas Series
                for nombre, total in doctores.items():
                    rows_doc.append([
                        "" if str(nombre).lower() == "nan" else nombre,
                        total
                    ])

            self._zebra_insert_all(self.tree_res_doc, rows_doc)

            # =============================================================
            # 2) RESUMEN: PROMOTORES
            # =============================================================
            rows_pro = []
            promotores = resumen.get("promotores", {})

            if isinstance(promotores, dict):
                for nombre, total in promotores.items():
                    rows_pro.append([
                        "" if str(nombre).lower() == "nan" else nombre,
                        total
                    ])
            elif hasattr(promotores, "items"):   # pandas Series
                for nombre, total in promotores.items():
                    rows_pro.append([
                        "" if str(nombre).lower() == "nan" else nombre,
                        total
                    ])

            self._zebra_insert_all(self.tree_res_pro, rows_pro)

            # =============================================================
            # 3) RESUMEN: TÉCNICOS
            # =============================================================
            rows_tec = []
            tecnicos = resumen.get("tecnicos", {})

            if isinstance(tecnicos, dict):
                for nombre, total in tecnicos.items():
                    rows_tec.append([
                        "" if str(nombre).lower() == "nan" else nombre,
                        total
                    ])
            elif hasattr(tecnicos, "items"):   # pandas Series
                for nombre, total in tecnicos.items():
                    rows_tec.append([
                        "" if str(nombre).lower() == "nan" else nombre,
                        total
                    ])

            self._zebra_insert_all(self.tree_res_tec, rows_tec)

            # =============================================================
            # 4) DETALLE COMPLETO DEL CORTE
            # =============================================================
            rows_det = []

            for _, r in df.iterrows():
                trat = r.get("tratamiento", "")
                trat = trat.split(" - ", 1)[-1]

                vals = [
                    r.get("fecha", ""),
                    r.get("paciente", ""),
                    trat,
                    r.get("doc_nombre", "") or "",
                    r.get("pro_nombre", "") or "",
                    r.get("tec_nombre", "") or "",
                    r.get("subtotal", ""),
                    r.get("precio_unitario", ""),
                    r.get("cantidad", ""),
                    r.get("comision_doctor", ""),
                    r.get("comision_promotor", ""),
                    r.get("utilidad_fila", "")
                ]

                # Limpiar valores tipo nan
                vals = [("" if str(v).lower() == "nan" else v) for v in vals]
                rows_det.append(vals)

            self._zebra_insert_all(self.tree_detalle, rows_det)

        except Exception as e:
            messagebox.showerror("Cortes", f"Error al cargar detalle:\n{e}")

    # -------------------------
    # KPIs
    # -------------------------
    """def _build_kpis_tab(self):
        frm = self.tab_kpis
        top = ttk.Frame(frm); top.pack(fill="x", padx=12, pady=8)
        self._kpi_desde = tk.StringVar(); self._kpi_hasta = tk.StringVar(); self._kpi_pdf = tk.BooleanVar(value=True)
        ttk.Label(top, text="Desde").pack(side="left"); de_kpi_desde = self._mk_date_entry(top, var=self._kpi_desde); de_kpi_desde.pack(side="left", padx=6)
        ttk.Label(top, text="Hasta").pack(side="left"); de_kpi_hasta = self._mk_date_entry(top, var=self._kpi_hasta); de_kpi_hasta.pack(side="left", padx=6)
        ttk.Checkbutton(top, text="PDF", variable=self._kpi_pdf).pack(side="left", padx=10); ttk.Button(top, text="Calcular", command=self._do_kpis).pack(side="left", padx=8)
        sep = ttk.Separator(frm, orient="horizontal"); sep.pack(fill="x", padx=12, pady=8)
        body = ttk.Frame(frm); body.pack(fill="both", expand=True, padx=12, pady=8)
        left = ttk.Frame(body); left.pack(side="left", fill="y"); right = ttk.Frame(body); right.pack(side="left", fill="both", expand=True, padx=18)
        self.var_ingresos = tk.StringVar(value="$0.00"); self.var_pacientes = tk.StringVar(value="0"); self.var_ticket = tk.StringVar(value="$0.00"); self.var_retencion = tk.StringVar(value="0.0%")
        def row_stat(parent,label,var): r=ttk.Frame(parent); r.pack(anchor="w", pady=4); ttk.Label(r,text=label+": ", width=18, anchor="e").pack(side="left"); ttk.Label(r,textvariable=var,foreground="#222").pack(side="left")
        ttk.Label(left, text="Resumen", font=("Segoe UI",11,"bold")).pack(anchor="w", pady=(0,8))
        row_stat(left,"Ingresos",self.var_ingresos); row_stat(left,"Pacientes únicos",self.var_pacientes); row_stat(left,"Ticket promedio",self.var_ticket); row_stat(left,"Retención",self.var_retencion)
        self._tops_text = tk.Text(right, height=22); self._tops_text.pack(fill="both", expand=True)"""
    
    def _build_kpis_tab(self):
        frm = self.tab_kpis

        # ---- FILA SUPERIOR: FECHAS + BOTÓN ----
        top = ttk.Frame(frm); top.pack(fill="x", padx=12, pady=8)

        self._kpi_desde = tk.StringVar()
        self._kpi_hasta = tk.StringVar()
        self._kpi_pdf = tk.BooleanVar(value=True)

        de_kpi_desde = self._mk_date_selector(top, var=self._kpi_desde)
        de_kpi_desde.pack(side="left", padx=6)

        de_kpi_hasta = self._mk_date_selector(top, var=self._kpi_hasta)
        de_kpi_hasta.pack(side="left", padx=6)

        ttk.Checkbutton(top, text="PDF", variable=self._kpi_pdf).pack(side="left", padx=10)
        ttk.Button(top, text="Calcular KPIs", command=self._do_kpis).pack(side="left", padx=8)

        # ---- SELECTOR DE PERIODO ----
        periodo_box = ttk.Frame(frm); periodo_box.pack(fill="x", padx=12, pady=6)
        ttk.Label(periodo_box, text="Ver por:").pack(side="left")
        self._kpi_periodo = tk.StringVar(value="Día")
        cb = ttk.Combobox(periodo_box, textvariable=self._kpi_periodo,
                        values=["Día", "Semana", "Mes"],
                        width=10, state="readonly")
        cb.pack(side="left", padx=6)
        ttk.Button(periodo_box, text="Actualizar historial",
                command=self._actualizar_historial).pack(side="left", padx=6)

        ttk.Separator(frm, orient="horizontal").pack(fill="x", padx=12, pady=8)

        # ---- BODY ----
        body = ttk.Frame(frm); body.pack(fill="both", expand=True, padx=12, pady=8)
        left = ttk.Frame(body); left.pack(side="left", fill="y")
        right = ttk.Frame(body); right.pack(side="left", fill="both", expand=True, padx=18)

        # -----------------------------
        # PANEL IZQUIERDO — RESUMEN
        # -----------------------------
        self.var_ingresos = tk.StringVar(value="$0.00")
        self.var_utilidad = tk.StringVar(value="$0.00")
        self.var_pacientes = tk.StringVar(value="0")
        self.var_ticket = tk.StringVar(value="$0.00")
        self.var_retencion = tk.StringVar(value="0.0%")

        def row_stat(parent, label, var):
            r = ttk.Frame(parent); r.pack(anchor="w", pady=4)
            ttk.Label(r, text=label + ": ", width=18, anchor="e").pack(side="left")
            ttk.Label(r, textvariable=var, foreground="#222").pack(side="left")

        ttk.Label(left, text="Resumen", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0,8))

        row_stat(left, "Ingresos", self.var_ingresos)
        row_stat(left, "Utilidad", self.var_utilidad)
        row_stat(left, "Pacientes únicos", self.var_pacientes)
        row_stat(left, "Ticket promedio", self.var_ticket)
        row_stat(left, "Retención", self.var_retencion)

        # -----------------------------
        # PANEL DERECHO — TOP LISTS
        # -----------------------------
        self._tops_text = tk.Text(right, height=12)
        self._tops_text.pack(fill="both", expand=False)

        # -----------------------------
        # HISTORIAL — GRÁFICO + TABLA
        # -----------------------------
        hist_frame = ttk.LabelFrame(right, text="Historial de ventas")
        hist_frame.pack(fill="both", expand=True, pady=(8,0))

        # Matplotlib
        fig = plt.Figure(figsize=(5,2.5))
        self._kpi_fig = fig
        self._kpi_ax = fig.add_subplot(111)
        self._kpi_canvas = FigureCanvasTkAgg(fig, master=hist_frame)
        self._kpi_canvas.get_tk_widget().pack(fill="both", expand=True)

        # Tabla
        tv_frame = ttk.Frame(hist_frame); tv_frame.pack(fill="both", expand=True, pady=(6,0))
        cols = ("fecha", "ventas")
        self.tree_kpi_hist = ttk.Treeview(tv_frame, columns=cols, show="headings", height=8)
        self.tree_kpi_hist.heading("fecha", text="Fecha")
        self.tree_kpi_hist.heading("ventas", text="Ventas")
        self.tree_kpi_hist.column("fecha", width=160, anchor="center")
        self.tree_kpi_hist.column("ventas", width=120, anchor="e")

        self.tree_kpi_hist = self._wrap_tree_with_scroll(tv_frame, self.tree_kpi_hist)
        self._apply_zebra(self.tree_kpi_hist)


    def _do_kpis(self):
        if not self._ensure_repo():
            return

        d = self._kpi_desde.get().strip()
        h = self._kpi_hasta.get().strip()
        if not d or not h:
            messagebox.showwarning("KPIs", "Indica fechas")
            return

        # -------------------------
        # Normalizador de tops
        # -------------------------
        def normalize_top(obj):
            """Convierte cualquier tipo (dict, Series, DataFrame, lista) → lista de tuplas (k,v)."""
            try:
                import pandas as _pd
            except:
                _pd = None

            if obj is None:
                return []

            if isinstance(obj, dict):
                return list(obj.items())

            if _pd is not None and isinstance(obj, _pd.Series):
                return list(obj.items())

            if _pd is not None and isinstance(obj, _pd.DataFrame):
                rows = []
                for _, r in obj.iterrows():
                    vals = list(r.values)
                    if len(vals) >= 2:
                        rows.append((str(vals[0]), vals[1]))
                return rows

            if isinstance(obj, (list, tuple)):
                out = []
                for it in obj:
                    if isinstance(it, (list, tuple)) and len(it) >= 2:
                        out.append((str(it[0]), it[1]))
                    else:
                        out.append((str(it), None))
                return out

            return [(str(obj), None)]

        # -------------------------
        # Formateo de TOPs
        # -------------------------
        def top_to_text(title, obj, limit=5):
            lst = normalize_top(obj)
            if not lst:
                return f"{title}\n  — (sin datos)\n\n"

            lines = [title]
            for i, (k, v) in enumerate(lst[:limit], start=1):
                if v is None:
                    lines.append(f"  {i}. {k}")
                else:
                    try:
                        vnum = float(v)
                        lines.append(f"  {i}. {k} — {vnum:.2f}")
                    except:
                        lines.append(f"  {i}. {k} — {v}")
            lines.append("")
            return "\n".join(lines)

        try:
            # ============================================================
            # CALCULAR KPIs (backend ya corregido)
            # ============================================================
            kpis = KPIService(self.repo, self.cfg).calcular(_parse_date(d), _parse_date(h))

            if not isinstance(kpis, dict):
                kpis = dict(kpis)

            # ---------------------------------------
            # Lectura segura de métricas principales
            # ---------------------------------------
            ingresos = float(kpis.get("ingresos_periodo", 0.0))
            utilidad = float(kpis.get("utilidad_periodo", 0.0))
            pacientes = int(kpis.get("pacientes_unicos", 0))
            ticket = float(kpis.get("ticket_promedio", 0.0))
            retencion = float(kpis.get("retencion", 0.0))

            # ---------------------------------------
            # Actualizar variables de UI
            # ---------------------------------------
            self.var_ingresos.set(f"${ingresos:,.2f}")
            self.var_utilidad.set(f"${utilidad:,.2f}")
            self.var_pacientes.set(str(pacientes))
            self.var_ticket.set(f"${ticket:,.2f}")
            self.var_retencion.set(f"{retencion*100:.1f}%")

            # ---------------------------------------
            # Construcción texto de TOPs
            # ---------------------------------------
            txt_parts = []

            txt_parts.append(
                top_to_text("Top tratamientos (ingresos)", kpis.get("top_tratamientos_ingresos"))
            )
            txt_parts.append(
                top_to_text("Top tratamientos (volumen)", kpis.get("top_tratamientos_volumen"))
            )
            txt_parts.append(
                top_to_text("Top dentistas (ingresos)", kpis.get("top_dentistas_ingresos"))
            )
            txt_parts.append(
                top_to_text("Top dentistas (volumen)", kpis.get("top_dentistas_volumen"))
            )

            # Extra tops
            for k in kpis:
                if k.startswith("top_") and k not in (
                    "top_tratamientos_ingresos",
                    "top_tratamientos_volumen",
                    "top_dentistas_ingresos",
                    "top_dentistas_volumen",
                ):
                    txt_parts.append(top_to_text(k.replace("_", " ").capitalize(), kpis[k]))

            # Render final
            self._tops_text.delete("1.0", "end")
            self._tops_text.insert("end", "\n".join(txt_parts))

            # ---------------------------------------
            # Actualizar historial del gráfico
            # ---------------------------------------
            try:
                self._actualizar_historial()
            except Exception as e:
                print("No pude actualizar historial:", e)

            # ---------------------------------------
            # PDF opcional
            # ---------------------------------------
            if self._kpi_pdf.get():
                try:
                    from consultorio_items import export_pdf_kpis
                    path = export_pdf_kpis(self.cfg, kpis)
                    messagebox.showinfo("KPIs", f"PDF generado:\n{path}")
                except Exception as e:
                    messagebox.showwarning("KPIs - PDF", f"No pude generar PDF:\n{e}")

        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            messagebox.showerror("KPIs - Error", f"Error al calcular KPIs:\n{e}\n\n{tb}")
            print("Error en KPIs:", e)
            print(tb)


    def _group_sales(self, df_ventas: "pd.DataFrame"):
        """
        Recibe DataFrame con columnas ['fecha','ventas'] (fecha datetime) y devuelve
        dos listas (x_labels, y_values) listas para graficar / mostrar.
        """
        # Normalizar formatos
        df = df_ventas.copy()
        if df.empty:
            return [], []
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        df = df.dropna(subset=["fecha"])
        # etiquetas legibles
        x = [d.strftime("%Y-%m-%d") for d in df["fecha"].dt.to_pydatetime()]
        y = [float(v) for v in df["ventas"].tolist()]
        return x, y

    def _actualizar_historial(self):
        """Lee fechas y periodo desde los widgets y pide datos al backend KPIService, luego actualiza gráfica y tabla."""
        if not self._ensure_repo(): return
        d = self._kpi_desde.get().strip(); h = self._kpi_hasta.get().strip()
        if not d or not h:
            messagebox.showwarning("KPIs","Indica fechas para el historial")
            return
        try:
            desde = _parse_date(d); hasta = _parse_date(h)
        except Exception:
            messagebox.showwarning("KPIs","Fechas inválidas")
            return

        periodo_map = {"Día": "dia", "Semana": "semana", "Mes": "mes"}
        periodo_sel = self._kpi_periodo.get() or "Día"
        periodo = periodo_map.get(periodo_sel, "dia")

        ksvc = KPIService(self.repo, self.cfg)
        ventas_df = ksvc.ventas_totales_por_periodo(desde, hasta, periodo=periodo)
        # actualizar tabla
        self.tree_kpi_hist.delete(*self.tree_kpi_hist.get_children())
        if ventas_df.empty:
            messagebox.showinfo("Historial", "No hay ventas en el rango seleccionado.")
            # limpiar gráfica
            self._kpi_ax.clear()
            self._kpi_ax.set_title("Ventas")
            self._kpi_canvas.draw()
            return

        # llenar tabla
        for _, r in ventas_df.iterrows():
            fecha_txt = pd.to_datetime(r["fecha"]).strftime("%Y-%m-%d")
            self.tree_kpi_hist.insert("", "end", values=(fecha_txt, f"{float(r['ventas']):.2f}"))

        # graficar
        x, y = self._group_sales(ventas_df)
        self._kpi_ax.clear()
        self._kpi_ax.plot(x, y, marker="o", linewidth=1)
        self._kpi_ax.set_title(f"Ventas totales por {periodo_sel}")
        # rotar etiquetas si muchas
        for label in self._kpi_ax.get_xticklabels():
            label.set_rotation(40)
            label.set_ha("right")
        self._kpi_ax.set_ylabel("Ventas")
        self._kpi_fig.tight_layout()
        try:
            self._kpi_canvas.draw()
        except Exception:
            pass

    def _format_top_block(self, titulo: str, data: dict, pct: dict = None) -> str:
        """
        Devuelve un bloque de texto bonito:
            Top tratamientos (ingresos)
            1. Resina — 4350.00 (23%)
            2. Limpieza — 3100.00 (17%)

        pct puede ser None.
        """
        lines = [titulo]

        if not data:
            lines.append("   — (sin datos)")
            return "\n".join(lines) + "\n"

        for i, (k, v) in enumerate(data.items(), start=1):
            nombre = k if k.strip() else "(desconocido)"
            monto = f"{v:,.2f}"

            if pct and k in pct:
                porcentaje = f"{pct[k]}%"
                lines.append(f"   {i}. {nombre} — {monto}  ({porcentaje})")
            else:
                lines.append(f"   {i}. {nombre} — {monto}")

        return "\n".join(lines) + "\n"
    
    def _render_kpis_text(self, kpi):
        """
        Recibe el dict resultante de calcular() y arma el texto
        con los tops en formato profesional.
        """

        txt = ""

        # Top tratamientos (ingresos)
        txt += self._format_top_block(
            "Top tratamientos (ingresos)",
            kpi.get("top_tratamientos_ingresos", {})
        )

        # Top tratamientos (volumen)
        txt += self._format_top_block(
            "Top tratamientos (volumen)",
            kpi.get("top_tratamientos_volumen", {})
        )

        # Top dentistas ingresos
        txt += self._format_top_block(
            "Top dentistas (ingresos)",
            kpi.get("top_dentistas_ingresos", {})
        )

        # Top dentistas volumen + porcentaje
        txt += self._format_top_block(
            "Top dentistas (volumen)",
            kpi.get("top_dentistas_volumen", {}),
            pct=kpi.get("top_dentistas_porcentaje", {})
        )

        return txt



######################################################################################################################################
         #############COMISIONES###################################
    ###############################################

    # ============================================================
#  PESTAÑA DE COMISIONES — FRONTEND COMPLETO Y FUNCIONAL
# ============================================================

    def _build_comisiones_tab(self):
        
        frm = self.tab_comisiones
        

        # ============================================================
        #  BLOQUE 1 — RESUMEN DEL DÍA POR DOCTOR
        # ============================================================

        frame_resumen = ttk.LabelFrame(frm, text="Resumen del día por Doctor")
        frame_resumen.pack(fill="x", padx=12, pady=8)

        frm_top = ttk.Frame(frame_resumen)
        frm_top.pack(fill="x", pady=5)

        # --- Doctor ---
        ttk.Label(frm_top, text="Doctor:").pack(side="left", padx=(0, 5))
        self.cmb_doc_comisiones = ttk.Combobox(frm_top, state="readonly", width=22)
        self.cmb_doc_comisiones.pack(side="left", padx=(0, 15))

        # --- Fecha ---
        ttk.Label(frm_top, text="Fecha:").pack(side="left", padx=(0, 5))
        self.date_fecha_comisiones = DateEntry(
            frm_top,
            width=12,
            date_pattern="yyyy-mm-dd",
            state="readonly"
        )
        self.date_fecha_comisiones.pack(side="left", padx=(0, 15))

        # --- Botón Cargar ---
        ttk.Button(
            frm_top,
            text="Cargar",
            command=self._cargar_comisiones_dia
        ).pack(side="left")

        # --- Tabla Resumen ---
        cols1 = ("hora", "paciente", "total", "metodo", "moneda", "comision", "pagado", "pendiente")
        self.tree_com_resumen = ttk.Treeview(frame_resumen, columns=cols1, show="headings", height=10)

        for c in cols1:
            self.tree_com_resumen.heading(c, text=c.capitalize())
            self.tree_com_resumen.column(c, width=120, anchor="center")

        self.tree_com_resumen.pack(fill="x", padx=5, pady=5)

        # --- Totales ---
        frm_tot = ttk.Frame(frame_resumen)
        frm_tot.pack(anchor="w", padx=8, pady=5)

        ttk.Label(frm_tot, text="Total generado:").grid(row=0, column=0, sticky="w")
        self.lbl_com_gen = ttk.Label(frm_tot, text="0.00")
        self.lbl_com_gen.grid(row=0, column=1, padx=10)

        ttk.Label(frm_tot, text="Pagado:").grid(row=0, column=2, sticky="w")
        self.lbl_com_pag = ttk.Label(frm_tot, text="0.00")
        self.lbl_com_pag.grid(row=0, column=3, padx=10)

        ttk.Label(frm_tot, text="Pendiente:").grid(row=0, column=4, sticky="w")
        self.lbl_com_pen = ttk.Label(frm_tot, text="0.00")
        self.lbl_com_pen.grid(row=0, column=5, padx=10)

        # ============================================================
        #  BLOQUE 2 — ABONOS A COMISIONES
        # ============================================================

        frame_abonos = ttk.LabelFrame(frm, text="Abonos a Comisiones")
        frame_abonos.pack(fill="x", padx=12, pady=8)

        frm_ab_top = ttk.Frame(frame_abonos)
        frm_ab_top.pack(fill="x", pady=5)

        ttk.Label(frm_ab_top, text="Personal:").pack(side="left", padx=(0, 5))
        self.cmb_abonos_personal = ttk.Combobox(frm_ab_top, state="readonly", width=22)
        self.cmb_abonos_personal.pack(side="left", padx=(0, 15))

        ttk.Button(
            frm_ab_top,
            text="Ver pendientes",
            command=self._cargar_pendientes_personal
        ).pack(side="left")

        cols2 = ("fecha", "paciente", "total", "comision", "abonado", "restante", "moneda", "metodo")
        self.tree_com_pendientes = ttk.Treeview(frame_abonos, columns=cols2, show="headings", height=8)

        for c in cols2:
            self.tree_com_pendientes.heading(c, text=c.capitalize())
            self.tree_com_pendientes.column(c, width=120, anchor="center")

        self.tree_com_pendientes.pack(fill="x", padx=5, pady=5)

        frm_ab_btn = ttk.Frame(frame_abonos)
        frm_ab_btn.pack(anchor="e", padx=8, pady=5)

        ttk.Button(
            frm_ab_btn,
            text="Registrar Abono",
            command=self._abrir_ventana_abono
        ).pack(side="right")

        # ============================================================
        #  BLOQUE 3 — HISTORIAL DE PAGOS
        # ============================================================

        frame_hist = ttk.LabelFrame(frm, text="Historial de Pagos de Comisiones")
        frame_hist.pack(fill="both", padx=12, pady=8, expand=True)

        cols3 = ("fecha_pago", "personal", "paciente", "consulta", "monto", "moneda", "nota")
        self.tree_com_hist = ttk.Treeview(frame_hist, columns=cols3, show="headings", height=8)

        for c in cols3:
            self.tree_com_hist.heading(c, text=c.capitalize())
            self.tree_com_hist.column(c, width=120, anchor="center")

        self.tree_com_hist.pack(fill="both", padx=5, pady=5, expand=True)
        self._cargar_personal_comisiones()
        self._cargar_historial_abonos()

    def _filtrar_personal_activo(self, df: pd.DataFrame) -> pd.DataFrame:
        """Devuelve únicamente filas marcadas como activas en la columna 'activo'."""
        if "activo" not in df.columns:
            return df

        def _is_active(val):
            try:
                return float(val) == 1.0
            except (TypeError, ValueError):
                pass
            s = str(val).strip().lower()
            return s in {"1", "true", "si", "sí", "activo", "yes", "y"}

        mask = df["activo"].apply(_is_active)
        return df[mask]

    def _cargar_personal_comisiones(self):
        if not self._ensure_repo():
            return

        df = self.repo.dfs["Personal"]
        df = self._filtrar_personal_activo(df)

        nombres = df["nombre"].tolist()

        self.cmb_doc_comisiones["values"] = nombres
        self.cmb_abonos_personal["values"] = nombres
        if nombres:
            self.cmb_doc_comisiones.set(nombres[0])
            self.cmb_abonos_personal.set(nombres[0])

    # ============================================================
    #  VENTANA EMERGENTE PARA CONFIRMAR ABONO
    # ============================================================

    def _abrir_ventana_abono(self):
        import tkinter as tk
        from tkinter import ttk, messagebox

        sel = self.tree_com_pendientes.focus()
        if not sel:
            messagebox.showwarning("Abono", "Selecciona una comisión pendiente.")
            return

        data = self.tree_com_pendientes.item(sel)["values"]
        print("[DEBUG] Fila seleccionada (raw):", self.tree_com_pendientes.item(sel))
        if not data:
            return

        fecha, paciente, total, comi, abonado, restante, moneda, metodo = data

        win = tk.Toplevel(self)
        win.title("Confirmar Abono")
        win.geometry("350x300")
        win.grab_set()

        ttk.Label(win, text=f"Personal: {self.cmb_abonos_personal.get()}",
                font=("Arial", 10, "bold")).pack(pady=5)

        ttk.Label(win, text=f"Paciente: {paciente}").pack()
        ttk.Label(win, text=f"Moneda: {moneda}").pack()
        ttk.Label(win, text=f"Comisión restante: ${restante}").pack(pady=5)

        ttk.Label(win, text="Monto a abonar:").pack(pady=3)
        ent_monto = ttk.Entry(win)
        ent_monto.pack()

        ttk.Label(win, text="Nota (opcional):").pack(pady=3)
        txt_nota = tk.Text(win, height=4, width=30)
        txt_nota.pack()

        def confirmar():
            monto = ent_monto.get().strip()
            nota = txt_nota.get("1.0", "end").strip()

            if not monto:
                messagebox.showwarning("Abono", "Indica un monto.")
                return

            self._registrar_abono(
                personal=self.cmb_abonos_personal.get(),
                fecha=fecha,
                paciente=paciente,
                monto=monto,
                nota=nota
            )

            win.destroy()

        ttk.Button(win, text="Confirmar Abono", command=confirmar).pack(pady=10)


    # ============================================================
    # MÉTODOS DEL FRONTEND (se conectarán al backend)
    # ============================================================

    def _cargar_comisiones_dia(self):
        """
        1. Limpia tabla de resumen
        2. Llama al backend (cuando esté listo)
        3. Rellena árbol y totales
        """
        if not self._ensure_repo():
            return

        doctor = self.cmb_doc_comisiones.get().strip()
        fecha_sel = self.date_fecha_comisiones.get_date() if hasattr(self.date_fecha_comisiones, "get_date") else None

        if not doctor:
            messagebox.showwarning("Comisiones", "Selecciona un doctor.")
            return
        if not fecha_sel:
            messagebox.showwarning("Comisiones", "Selecciona una fecha.")
            return

        self.tree_com_resumen.delete(*self.tree_com_resumen.get_children())

        try:
            df, tot = self._comision_service.resumen_dia_por_doctor(fecha_sel, doctor)
        except Exception as e:
            messagebox.showerror("Comisiones", f"Error cargando comisiones:\n{e}")
            return

        for _, r in df.iterrows():
            vals = (
                r.get("hora", ""),
                r.get("paciente", ""),
                f"{float(r.get('total', 0.0)):.2f}",
                r.get("metodo", ""),
                r.get("moneda", ""),
                f"{float(r.get('comision', 0.0)):.2f}",
                f"{float(r.get('pagado', 0.0)):.2f}",
                f"{float(r.get('pendiente', 0.0)):.2f}",
            )
            self.tree_com_resumen.insert("", "end", values=vals)

        self.lbl_com_gen.config(text=f"{float(tot.get('total_generado', 0.0)):.2f}")
        self.lbl_com_pag.config(text=f"{float(tot.get('pagado', 0.0)):.2f}")
        self.lbl_com_pen.config(text=f"{float(tot.get('pendiente', 0.0)):.2f}")
        self._cargar_historial_abonos()


    def _cargar_pendientes_personal(self):
        """
        Carga el listado de comisiones pendientes del personal seleccionado.
        """
        if not self._ensure_repo():
            return

        nombre = self.cmb_abonos_personal.get().strip()
        if not nombre:
            messagebox.showwarning("Comisiones", "Selecciona personal.")
            return

        personal = self.repo.dfs.get("Personal", pd.DataFrame()).copy()
        personal["nombre_norm"] = personal.get("nombre", "").fillna("").apply(lambda s: str(s).strip().lower())
        match = personal[personal["nombre_norm"] == nombre.strip().lower()]
        if match.empty:
            match = personal[personal.get("nombre", "").fillna("").str.contains(nombre, case=False, na=False)]
        if match.empty:
            messagebox.showwarning("Comisiones", "No encontré al personal seleccionado.")
            return

        id_personal = str(match.iloc[0].get("id_personal", ""))
        self.tree_com_pendientes.delete(*self.tree_com_pendientes.get_children())

        try:
            df = self._comision_service.pendientes_por_personal(id_personal)
        except Exception as e:
            messagebox.showerror("Comisiones", f"Error cargando pendientes:\n{e}")
            return

        for _, r in df.iterrows():
            vals = (
                r.get("fecha", ""),
                r.get("paciente", ""),
                f"{float(r.get('total', 0.0)):.2f}",
                f"{float(r.get('comision', 0.0)):.2f}",
                f"{float(r.get('abonado', 0.0)):.2f}",
                f"{float(r.get('restante', 0.0)):.2f}",
                r.get("moneda", ""),
                r.get("metodo", ""),
            )
            self.tree_com_pendientes.insert("", "end", values=vals)
        self._cargar_historial_abonos()


    def _registrar_abono(self, personal, fecha, paciente, monto, nota):
        """
        Registra un abono en backend.
        """
        if not self._ensure_repo():
            return

        try:
            monto_val = float(monto)
        except Exception:
            messagebox.showwarning("Abono", "El monto debe ser numérico.")
            return

        try:
            if hasattr(_parse_date, "__call__"):
                parsed_fecha = _parse_date(str(fecha))
                if isinstance(parsed_fecha, datetime):
                    fecha_dt = parsed_fecha.date()
                elif isinstance(parsed_fecha, date):
                    fecha_dt = parsed_fecha
                else:
                    fecha_dt = parsed_fecha
            else:
                fecha_dt = date.fromisoformat(str(fecha))
        except Exception as e:
            print("[DEBUG] No pude parsear fecha del TreeView:", fecha, "->", e)
            fecha_dt = None

        try:
            print("[DEBUG] Registrar abono -> personal:", personal,
                  "fecha_raw:", fecha,
                  "fecha_parseada:", fecha_dt or date.today(),
                  "paciente:", paciente,
                  "monto:", monto_val,
                  "nota:", nota)
            aid = self._comision_service.registrar_abono_por_nombre(
                personal_nombre=personal,
                fecha_consulta=fecha_dt or date.today(),
                paciente_nombre=paciente,
                monto=monto_val,
                nota=nota or ""
            )
        except Exception as e:
            messagebox.showerror("Abono", f"No pude registrar abono:\n{e}")
            return

        if aid:
            messagebox.showinfo("Abono", f"Abono registrado (ID {aid}).")
            self._cargar_pendientes_personal()
            self._cargar_comisiones_dia()
            self._cargar_historial_abonos()
        else:
            messagebox.showwarning("Abono", "No se encontró la consulta para registrar el abono.")

    def _cargar_historial_abonos(self):
        if not self._ensure_repo():
            return

        if not self._comision_service:
            return

        self.tree_com_hist.delete(*self.tree_com_hist.get_children())

        try:
            df = self._comision_service.historial_abonos()
        except Exception as e:
            messagebox.showerror("Comisiones", f"Error cargando historial:\n{e}")
            return

        for _, r in df.iterrows():
            vals = (
                r.get("fecha_pago", ""),
                r.get("personal", ""),
                r.get("paciente", ""),
                r.get("consulta", ""),
                f"{float(r.get('monto', 0.0)):.2f}",
                r.get("moneda", ""),
                r.get("nota", ""),
            )
            self.tree_com_hist.insert("", "end", values=vals)

 


if __name__ == "__main__":
    info_licencia = evaluar_licencia()

    estado = info_licencia.get("estado") if isinstance(info_licencia, dict) else None

    if estado == "PRUEBA_VENCIDA":
        messagebox.showerror(
            "Licencia",
            "El periodo de prueba ha finalizado. Por favor, contacta al proveedor para activar la licencia.",
        )
        sys.exit()

    if estado == "PRUEBA_ACTIVA":
        app = ConsultorioGUI(dias_licencia_restantes=info_licencia.get("dias_restantes"))
    else:
        app = ConsultorioGUI()

    app.mainloop()
