# Dashboard Quibdó - Configuración de Datos

##  PROBLEMA SOLUCIONADO: ModuleNotFoundError

El error ModuleNotFoundError ha sido corregido con:

1. **requirements.txt añadido** con todas las dependencias:
   - streamlit>=1.28.0
   - pandas>=2.0.0
   - plotly>=5.15.0
   - python-docx>=0.8.11
   - numpy>=1.24.0
   - openpyxl>=3.1.0

2. **Rutas configuradas automáticamente** para funcionar tanto local como en Streamlit Cloud

##  Para usar en Streamlit Cloud

Coloca tus archivos de datos en la carpeta data/:
- data/SEGUNDA_ENTREGA_7_SEPTIEMBRE.docx
- data/Instrumento_Analisis_Financiero_Grupo_5.xlsx

##  Deployment

El dashboard detecta automáticamente el entorno y usa las rutas correctas.

##  Acceso

- **GitHub:** https://github.com/efrenbohorquez/reporte-quibdo
- **Rama:** dashboard-optimizado
- **Archivo principal:** dashboard_quibdo.py
