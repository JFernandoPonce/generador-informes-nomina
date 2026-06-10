# 📄 Generador de Informes de Nómina — Distrito 17D08

Aplicación web local desarrollada en Python/Flask que automatiza la generación de informes técnicos de novedades de nómina para instituciones del Ministerio de Educación del Ecuador.

Diseñada para eliminar trabajo manual repetitivo: compara dos distributivos Excel mensuales, detecta automáticamente altas y bajas de personal, clasifica al personal en 6 secciones según régimen y modalidad, y genera el informe final en Word (.docx) listo para firma.

---

## ✨ Funcionalidades

- **Comparación automática** de distributivos Excel entre mes anterior y mes en curso
- **Detección de novedades**: ingresos, salidas y cambios de personal
- **Clasificación automática** del personal en 6 categorías (LOSEP / Otros Regímenes / Código de Trabajo)
- **Gestión de novedades manuales**: encargos, jubilaciones, suspensiones, etc.
- **Generación de documentos**: informe Word con tablas formateadas + Excel de verificación
- Interfaz web paso a paso, 100% local, sin dependencia de internet en uso

---

## 🛠️ Stack

| Capa | Tecnología |
|---|---|
| Backend | Python 3.8+, Flask 3.0 |
| Procesamiento de datos | Pandas, openpyxl, xlrd |
| Generación de documentos | python-docx |
| Frontend | HTML/CSS (templates Jinja2) |

---

## 🚀 Instalación y uso

### Requisitos
- Python 3.8 o superior
- Conexión a internet solo para la primera instalación de dependencias

### Windows
```bash
# Doble clic en iniciar.bat
# Se abre automáticamente http://localhost:5000
```

### Linux / Mac
```bash
chmod +x iniciar.sh
./iniciar.sh
```

### Manual
```bash
pip install -r requirements.txt
python app.py
```

---

## 📋 Flujo de trabajo

```
Paso 1: Cargar archivos Excel (mes anterior + mes en curso)
         ↓
Paso 2: Completar datos generales del informe (fecha, responsable, destinatario)
         ↓
Paso 3: Revisar novedades detectadas automáticamente y ajustar motivos
         ↓
Paso 4: Generar y descargar informe Word + Excel de verificación
```

### Clasificación automática del personal

| Sección | Régimen | Modalidad | Enlace Presupuestario |
|---|---|---|---|
| 9.1 | LOSEP | Contratos Ocasionales | — |
| 9.2 | LOSEP | Nombramiento | — |
| 9.3 | Otros Regímenes | Contratos Ocasionales | 13, 14, 15, 19 |
| 9.4 | Otros Regímenes | Contratos Ocasionales | 6, 7, 8 |
| 9.5 | Otros Regímenes | Nombramiento | — |
| 9.6 | Código de Trabajo | Todos | — |

### Archivos generados
- `INFORME_XXX_MES_AÑO.docx` — Informe completo con tablas formateadas, listo para firma
- `TABLAS_XXX_MES_AÑO.xlsx` — Excel con cada tabla en hoja separada para verificación

---

## 💡 Contexto

Desarrollado para optimizar el proceso mensual de reporte de novedades de nómina del Distrito de Educación 17D08 (Ministerio de Educación, Ecuador). El proceso anterior requería construir manualmente las tablas del informe comparando los distributivos línea por línea — esta aplicación lo hace en segundos.

---

## 📁 Estructura del proyecto

```
generador-informes-nomina/
├── app.py                  # Lógica principal Flask + procesamiento de datos
├── requirements.txt        # Dependencias Python
├── Plantilla.docx          # Plantilla base del informe Word
├── iniciar.bat             # Script de arranque Windows
├── iniciar.sh              # Script de arranque Linux/Mac
├── templates/
│   └── index.html          # Interfaz web (wizard paso a paso)
└── static/                 # Assets estáticos
```

---

## 🔒 Privacidad

La aplicación es 100% local. Ningún dato de nómina es enviado a servidores externos. Todo el procesamiento ocurre en la máquina del usuario.

---

*Desarrollado por [Juan Fernando Ponce](https://github.com/JFernandoPonce)*
