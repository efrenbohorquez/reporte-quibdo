import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Rediseño Institucional Quibdó", layout="wide")

st.title(" Estudio Técnico de Rediseño Institucional")
st.markdown("### Municipio de Quibdó, Chocó")

st.sidebar.title(" Dashboard Quibdó")
seccion = st.sidebar.selectbox(
    "Selecciona una sección:",
    [" Resumen", " Financiero", " Estructura", " Análisis"]
)

if seccion == " Resumen":
    st.header(" Resumen Ejecutivo")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Cumplimiento Ley 617", "50.54%", "vs 80% límite")
    with col2:
        st.metric("Margen Disponible", ",285 M", "+29.5%")
    with col3:
        st.metric("Población Afrocolombiana", "96.91%", "130,000 hab")
    with col4:
        st.metric("Víctimas Conflicto", "93,907", "registradas")

elif seccion == " Financiero":
    st.header(" Análisis Financiero")
    
    data_fin = pd.DataFrame({
        'Año': [2022, 2023, 2024, '2025 (Proy.)'],
        'ICLD (Mill $)': [32147, 34891, 36225, 37841],
        'Gastos Funcionamiento (Mill $)': [15286, 16789, 17894, 19108],
        '% ICLD Comprometido': [47.5, 48.1, 49.4, 50.5],
        'Margen Disponible (Mill $)': [10431, 11124, 11086, 11285]
    })
    
    st.dataframe(data_fin, use_container_width=True)
    
    fig = px.line(data_fin, x='Año', y='% ICLD Comprometido', 
                 markers=True, title='% ICLD Comprometido por Año')
    fig.add_hline(y=80, line_dash="dash", line_color="red")
    st.plotly_chart(fig, use_container_width=True)

elif seccion == " Estructura":
    st.header(" Estructura Organizacional")
    
    data_org = pd.DataFrame({
        'Dependencia': ['Secretarías', 'Coordinaciones', 'Asesores'],
        'Actual': [14, 9, 4],
        'Propuesta': [3, 6, 3]
    })
    
    fig = px.bar(data_org, x='Dependencia', y=['Actual', 'Propuesta'],
                title='Estructura Actual vs Propuesta', barmode='group')
    st.plotly_chart(fig, use_container_width=True)

elif seccion == " Análisis":
    st.header(" Brechas y Fortalezas")
    
    brechas = pd.DataFrame({
        'Área': ['Gobierno Digital', 'Servicios Públicos', 'Planeación', 'Finanzas', 'Ambiente'],
        'Brecha %': [71, 52, 45, 38, 33]
    })
    
    fig = px.bar(brechas, x='Brecha %', y='Área', orientation='h',
                title='Brechas Críticas por Área (%)', color='Brecha %')
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.markdown("**Autores:** John Miller Beltrán, Efrén Bohorquez, Edwin Suárez, Alexandra Calderón, Andrea Fuentes, Ernesto Romero, Daniela Gómez")
st.markdown("**Docente:** Rafael Arturo Amaya Mejía | **ESAP** - Bogotá D.C., 2025")
