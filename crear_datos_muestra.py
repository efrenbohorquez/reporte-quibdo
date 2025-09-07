# Archivo de datos de muestra para Streamlit Cloud
import pandas as pd
import numpy as np
from datetime import datetime

# Crear datos de ejemplo para el análisis financiero
np.random.seed(42)

# Datos principales del municipio
datos_municipio = {
    'Año': [2023, 2024, 2025],
    'Población': [120000, 122000, 124000],
    'Ingresos_Propios': [15000000, 16500000, 18000000],  # En millones
    'Transferencias': [45000000, 48000000, 51000000],    # En millones
    'Gastos_Totales': [58000000, 62000000, 65000000],    # En millones
    'Deuda_Publica': [12000000, 11500000, 11000000],     # En millones
    'Empleados': [450, 465, 480]
}

df_principal = pd.DataFrame(datos_municipio)

# Crear datos de estructura organizacional
dependencias = [
    'Alcaldía', 'Secretaría General', 'Secretaría de Hacienda', 
    'Secretaría de Planeación', 'Secretaría de Gobierno',
    'Secretaría de Educación', 'Secretaría de Salud',
    'Secretaría de Infraestructura', 'Secretaría de Ambiente'
]

estructura_org = {
    'Dependencia': dependencias,
    'Empleados': np.random.randint(5, 50, len(dependencias)),
    'Presupuesto_2025': np.random.randint(100000, 500000, len(dependencias)),  # En millones
    'Eficiencia': np.random.uniform(0.7, 0.95, len(dependencias)).round(2)
}

df_estructura = pd.DataFrame(estructura_org)

# Guardar como Excel con múltiples hojas
with pd.ExcelWriter('data/Instrumento_Analisis_Financiero_Grupo_5.xlsx') as writer:
    df_principal.to_excel(writer, sheet_name='Datos_Principales', index=False)
    df_estructura.to_excel(writer, sheet_name='Estructura_Organizacional', index=False)
    
    # Hoja adicional con indicadores
    indicadores = {
        'Indicador': ['Cobertura Educación', 'Cobertura Salud', 'Índice Desarrollo', 'Tasa Desempleo'],
        'Valor_2023': [85.5, 78.2, 0.65, 12.5],
        'Valor_2024': [87.1, 80.5, 0.68, 11.8],
        'Meta_2025': [90.0, 85.0, 0.72, 10.0]
    }
    df_indicadores = pd.DataFrame(indicadores)
    df_indicadores.to_excel(writer, sheet_name='Indicadores_Gestion', index=False)

print(' Archivo Excel de muestra creado exitosamente')
