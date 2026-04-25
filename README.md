# 🦷 App Consultorio Light Blue

Sistema de gestión clínica local para consultorios dentales. Desarrollado en Python con interfaz gráfica, orientado a automatizar los procesos administrativos y financieros del día a día de un consultorio.

---

## ¿Qué problema resuelve?

Los consultorios dentales pequeños y medianos suelen manejar citas, cobros, expedientes y comisiones de forma manual o en hojas de Excel desconectadas entre sí. Esta aplicación centraliza todo en un solo sistema local: desde el registro del paciente hasta el corte diario en PDF.

---

## ✅ Funcionalidades principales

### 📅 Gestión de citas
- Registro y seguimiento de citas por paciente y doctor
- Filtrado de consultas por doctor o por paciente
- Visualización del historial de atención

### 💳 Cobro y registro de pagos
- Registro de pago por consulta
- Soporte para múltiples formas de pago:
  - Efectivo (pesos mexicanos y dólares)
  - Tarjeta
- Separación automática por tipo de pago en reportes

### 🗂️ Expediente clínico
- Generación y almacenamiento de expediente clínico por paciente
- Registro guardado en formato Excel para fácil acceso y portabilidad

### 📊 Reportes y corte diario
- Generación de corte diario automático en PDF
- Resumen visual en pantalla con:
  - Total de consultas del día
  - Desglose por doctor
  - Separación de comisiones por doctor y por promotor

### 💰 Control de comisiones
- Cálculo automático de comisiones por doctor y promotor
- Registro de pagos de comisiones realizados
- Historial de saldo pendiente vs. pagado por cada colaborador

### 📈 Análisis financiero
- Consulta de ingresos totales por rango de fechas
- Producción generada por doctor (selección de periodo personalizado)
- Separación de ingresos por tipo de pago (efectivo MXN, dólares, tarjeta)

---

## 🛠️ Tecnologías utilizadas

| Área | Tecnología |
|---|---|
| Lenguaje | Python 3 |
| Interfaz gráfica | Tkinter + tkcalendar |
| Procesamiento de datos | pandas + numpy |
| Persistencia | Excel (.xlsx) mediante openpyxl |
| Gráficas | matplotlib |
| Generación de PDF | reportlab (pdfgen, platypus, lib) |

---

## 🚀 Instalación y uso

### Requisitos
- Python 3.9 o superior
- Dependencias listadas en `requirements.txt`

### Pasos
```bash
# 1. Clona el repositorio
git clone https://github.com/adan1219/app_consultorio_ligth_blue.git

# 2. Instala las dependencias
pip install -r requirements.txt

# 3. Ejecuta la aplicación
python consultorio_gui_items.py
```

> ⚠️ La aplicación es local — no requiere conexión a internet ni servidor externo.

---

## 📁 Estructura del proyecto

```
app_consultorio_light_blue/
├── consultorio_gui_items.py   # Interfaz gráfica (vistas y navegación)
├── consultorio_items.py       # Lógica de negocio y manejo de datos
├── licencia_local.py          # Control de licencia de uso local
├── consultorio_items.xlsx     # Archivo base de datos / expedientes
└── README.md
```

---

## 💡 Contexto de desarrollo

Este proyecto fue desarrollado para uso real en un consultorio dental, cubriendo las necesidades operativas del día a día:

- Reemplazó el registro manual en hojas de cálculo desconectadas
- Automatizó el cálculo y seguimiento de comisiones (proceso que antes tomaba tiempo considerable al cierre del día)
- Generó reportes PDF listos para entrega sin intervención manual adicional

---

## 👨‍💻 Autor

**Adán Padrón Salinas**  
Ingeniero Mecatrónico | Process Automation Developer  
📧 adan2812@hotmail.com  
🔗 [LinkedIn](https://linkedin.com/in/carlosramirez) <!-- actualiza este link -->

---

> Proyecto desarrollado con Python y asistencia de herramientas de IA (ChatGPT/Codex) para generación y revisión de código.
