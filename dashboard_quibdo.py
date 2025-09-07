"""
Dashboard de Rediseño Institucional - Municipio de Quibdó
Versión Optimizada y Modernizada

Mejoras implementadas:
- Arquitectura modular y orientada a objetos
- Manejo centralizado de configuración
- Carga optimizada de datos con cache
- Componentes UI reutilizables
- Mejor manejo de errores
- Documentación completa
- Tipos de datos con Type Hints
- Logging integrado
- CSS personalizado moderno
- Performance mejorado

Autor: Sistema de Análisis Institucional
Fecha: Septiembre 2025
Versión: 2.0
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
import numpy as np
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

# Configuración de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ===== CONFIGURACIÓN Y CONSTANTES =====

class ConfigPaths:
    """Configuración centralizada de rutas con detección dinámica"""
    # Rutas para desarrollo local
    WORD_DOC_LOCAL = Path(r"D:\Downloads\SEGUNDA ENTRGA 7 SEPTIEMBRE.docx")
    EXCEL_FILE_LOCAL = Path(r"D:\Downloads\Instrumento Análisis Financiero - Grupo 5.xlsx")

    # Rutas para Streamlit Cloud
    WORD_DOC_CLOUD = Path("data/SEGUNDA_ENTREGA_7_SEPTIEMBRE.docx")
    EXCEL_FILE_CLOUD = Path("data/Instrumento_Analisis_Financiero_Grupo_5.xlsx")

    @classmethod
    def get_word_doc_path(cls) -> Path:
        """Obtiene la ruta correcta del documento Word según el entorno"""
        # Primero intenta con la ruta de cloud (Streamlit Cloud)
        if cls.WORD_DOC_CLOUD.exists():
            logger.info(f"📁 Usando ruta Cloud para Word: {cls.WORD_DOC_CLOUD}")
            return cls.WORD_DOC_CLOUD

        # Si no existe, intenta con la ruta local
        if cls.WORD_DOC_LOCAL.exists():
            logger.info(f"📁 Usando ruta Local para Word: {cls.WORD_DOC_LOCAL}")
            return cls.WORD_DOC_LOCAL

        # Si ninguna existe, usar cloud por defecto (para mostrar mensaje de error controlado)
        logger.warning("⚠️ No se encontró archivo Word en rutas locales ni cloud")
        return cls.WORD_DOC_CLOUD

    @classmethod
    def get_excel_file_path(cls) -> Path:
        """Obtiene la ruta correcta del archivo Excel según el entorno"""
        # Primero intenta con la ruta de cloud (Streamlit Cloud)
        if cls.EXCEL_FILE_CLOUD.exists():
            logger.info(f"📁 Usando ruta Cloud para Excel: {cls.EXCEL_FILE_CLOUD}")
            return cls.EXCEL_FILE_CLOUD

        # Si no existe, intenta con la ruta local
        if cls.EXCEL_FILE_LOCAL.exists():
            logger.info(f"📁 Usando ruta Local para Excel: {cls.EXCEL_FILE_LOCAL}")
            return cls.EXCEL_FILE_LOCAL

        # Si ninguna existe, usar cloud por defecto (para mostrar mensaje de error controlado)
        logger.warning("⚠️ No se encontró archivo Excel en rutas locales ni cloud")
        return cls.EXCEL_FILE_CLOUD

    # Propiedades para compatibilidad con código existente
    @property
    def WORD_DOC(self) -> Path:
        return self.get_word_doc_path()

    @property
    def EXCEL_FILE(self) -> Path:
        return self.get_excel_file_path()

class AppConfig:
    """Configuración de la aplicación"""
    PAGE_TITLE = "Rediseño Institucional - Municipio de Quibdó"
    PAGE_ICON = "🏛️"
    LAYOUT = "wide"
    CACHE_TTL = 3600  # 1 hora en segundos

@dataclass
class DocumentAnalysis:
    """Estructura para análisis de documentos"""
    texto_completo: List[str]
    total_parrafos: int
    total_palabras: int
    total_tablas: int
    tablas: List[Dict]
    palabras_clave: Dict[str, int]

class SectionEnum(Enum):
    """Enumeración de secciones del dashboard"""
    DIAGNOSTICO = "🏠 Diagnóstico General"
    FINANCIERA = "📊 Situación Financiera"
    ESTRUCTURA = "🏛️ Estructura Organizacional"
    NORMATIVO = "⚖️ Cumplimiento Normativo"
    INDICADORES = "📈 Indicadores de Gestión"
    REDISENO = "🎯 Propuesta de Rediseño"
    TECNICA = "📋 Propuesta Técnica"
    CRONOGRAMA = "📅 Cronograma"
    PRESUPUESTO = "💰 Recursos y Presupuesto"

# Configuración de la página
st.set_page_config(
    page_title=AppConfig.PAGE_TITLE,
    page_icon=AppConfig.PAGE_ICON,
    layout=AppConfig.LAYOUT,
    initial_sidebar_state="expanded"
)

# ===== FUNCIONES DE DATOS OPTIMIZADAS =====

class DataLoader:
    """Clase para manejo centralizado y optimizado de carga de datos"""
    
    @staticmethod
    @st.cache_data(ttl=AppConfig.CACHE_TTL, show_spinner=True)
    def cargar_documento_word() -> Optional[DocumentAnalysis]:
        """
        Carga y analiza el documento Word de forma optimizada
        
        Returns:
            DocumentAnalysis o None si hay error
        """
        try:
            word_path = ConfigPaths.get_word_doc_path()
            if not word_path.exists():
                logger.warning(f"Archivo Word no encontrado: {word_path}")
                return None
            
            doc = Document(str(word_path))
            
            # Extraer texto optimizado
            texto_completo = [
                paragraph.text.strip() 
                for paragraph in doc.paragraphs 
                if paragraph.text.strip()
            ]
            
            if not texto_completo:
                logger.warning("Documento Word vacío o sin contenido válido")
                return None
            
            # Extraer tablas optimizado
            tablas_extraidas = []
            for i, table in enumerate(doc.tables):
                tabla_data = [
                    [cell.text.strip() for cell in row.cells]
                    for row in table.rows
                ]
                if tabla_data:  # Solo agregar tablas con contenido
                    tablas_extraidas.append({
                        'numero': i + 1,
                        'datos': tabla_data
                    })
            
            # Análisis de palabras clave optimizado
            texto_union = ' '.join(texto_completo).lower()
            palabras_clave = DataLoader._analizar_palabras_clave(texto_union)
            
            return DocumentAnalysis(
                texto_completo=texto_completo,
                total_parrafos=len(texto_completo),
                total_palabras=len(texto_union.split()),
                total_tablas=len(tablas_extraidas),
                tablas=tablas_extraidas,
                palabras_clave=palabras_clave
            )
            
        except Exception as e:
            logger.error(f"Error cargando documento Word: {e}")
            st.error(f"Error al cargar el documento Word: {e}")
            return None
    
    @staticmethod
    def _analizar_palabras_clave(texto: str) -> Dict[str, int]:
        """Análisis optimizado de palabras clave con categorización"""
        
        categorias_palabras = {
            'geograficas': ['quibdó', 'quibdo', 'municipio', 'chocó', 'choco', 'departamento', 'territorial'],
            'normativas': ['ley 617', 'ley617', 'decreto', 'resolución', 'normatividad', 'normativa', 'reglamento'],
            'organizacionales': ['estructura', 'organizacional', 'organigrama', 'cargo', 'dependencia', 'secretaría'],
            'gestion': ['gestión', 'calidad', 'control', 'seguimiento', 'indicador', 'proceso', 'procedimiento'],
            'financieras': ['financiero', 'presupuesto', 'fiscal', 'ingresos', 'gastos', 'recursos', 'costo'],
            'modernizacion': ['modernización', 'rediseño', 'reforma', 'transformación', 'innovación', 'mejoramiento']
        }
        
        resultado = {}
        
        # Análisis por categorías
        for categoria, palabras in categorias_palabras.items():
            total_categoria = sum(texto.count(palabra) for palabra in palabras)
            resultado[categoria] = total_categoria
            
            # Palabras individuales más importantes
            for palabra in palabras[:3]:
                resultado[palabra] = texto.count(palabra)
        
        return resultado
    
    @staticmethod
    @st.cache_data(ttl=AppConfig.CACHE_TTL, show_spinner=True)
    def cargar_datos_excel() -> Tuple[Optional[pd.DataFrame], Optional[Dict], Optional[List]]:
        """
        Carga optimizada de datos del archivo Excel
        
        Returns:
            Tupla con (datos_principales, todas_hojas, nombres_hojas)
        """
        try:
            excel_path = ConfigPaths.get_excel_file_path()
            if not excel_path.exists():
                logger.warning(f"Archivo Excel no encontrado: {excel_path}")
                return None, None, None
            
            excel_file = pd.ExcelFile(str(excel_path))
            hojas = excel_file.sheet_names
            
            if not hojas:
                logger.warning("Archivo Excel sin hojas válidas")
                return None, None, None
            
            # Datos principales (primera hoja)
            datos_principales = pd.read_excel(
                str(excel_path), 
                sheet_name=0,
                na_values=['', ' ', 'N/A', 'n/a', 'NULL']
            )
            
            # Todas las hojas con manejo de errores
            todas_las_hojas = {}
            for hoja in hojas:
                try:
                    df_hoja = pd.read_excel(
                        str(excel_path), 
                        sheet_name=hoja,
                        na_values=['', ' ', 'N/A', 'n/a', 'NULL']
                    )
                    
                    if not df_hoja.empty:
                        todas_las_hojas[hoja] = df_hoja
                    else:
                        logger.warning(f"Hoja '{hoja}' está vacía")
                        
                except Exception as e:
                    logger.warning(f"No se pudo leer la hoja '{hoja}': {e}")
            
            return datos_principales, todas_las_hojas, list(todas_las_hojas.keys())
            
        except Exception as e:
            logger.error(f"Error cargando archivo Excel: {e}")
            st.error(f"❌ Error al cargar el archivo Excel: {e}")
            return None, None, None

# ===== CSS PERSONALIZADO MODERNO =====

def load_custom_css():
    """Carga estilos CSS modernos y responsivos"""
    css_styles = """
    <style>
        /* Variables CSS para tema consistente */
        :root {
            --primary-color: #1e3c72;
            --secondary-color: #2a5298;
            --accent-color: #f0f2f6;
            --success-color: #28a745;
            --warning-color: #ffc107;
            --error-color: #dc3545;
            --text-color: #333333;
            --border-radius: 12px;
            --shadow: 0 4px 20px rgba(0,0,0,0.1);
            --transition: all 0.3s ease;
        }
        
        /* Header principal con gradiente */
        .main-header {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%);
            color: white;
            padding: 2rem;
            border-radius: var(--border-radius);
            text-align: center;
            margin-bottom: 2rem;
            box-shadow: var(--shadow);
        }
        
        /* Tarjetas de métricas mejoradas */
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: var(--border-radius);
            border-left: 4px solid var(--primary-color);
            margin-bottom: 1rem;
            transition: var(--transition);
            box-shadow: var(--shadow);
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        
        /* Alertas mejoradas */
        .alert-success {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border: 1px solid #c3e6cb;
            color: #155724;
            padding: 1.2rem;
            border-radius: var(--border-radius);
            border-left: 4px solid var(--success-color);
            margin: 1rem 0;
        }
        
        .alert-warning {
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 1.2rem;
            border-radius: var(--border-radius);
            border-left: 4px solid var(--warning-color);
            margin: 1rem 0;
        }
        
        .alert-error {
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
            border: 1px solid #f5c6cb;
            color: #721c24;
            padding: 1.2rem;
            border-radius: var(--border-radius);
            border-left: 4px solid var(--error-color);
            margin: 1rem 0;
        }
        
        /* Footer mejorado */
        .footer {
            text-align: center;
            color: #666;
            padding: 2rem 1rem;
            border-top: 1px solid #eee;
            margin-top: 3rem;
            background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
            border-radius: var(--border-radius);
        }
        
        /* Métricas personalizadas */
        [data-testid="metric-container"] {
            background: white;
            border: 1px solid #e0e0e0;
            padding: 1rem;
            border-radius: var(--border-radius);
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            transition: var(--transition);
        }
        
        [data-testid="metric-container"]:hover {
            box-shadow: var(--shadow);
            transform: translateY(-2px);
        }
    </style>
    """
    st.markdown(css_styles, unsafe_allow_html=True)

# ===== COMPONENTES UI MODERNOS =====

class UIComponents:
    """Componentes reutilizables y modernos de interfaz de usuario"""
    
    @staticmethod
    def create_info_box(title: str, content: str, box_type: str = "info"):
        """Crea una caja de información moderna con iconos"""
        icon_map = {
            "info": "ℹ️",
            "success": "✅",
            "warning": "⚠️",
            "error": "❌"
        }
        
        icon = icon_map.get(box_type, "ℹ️")
        
        st.markdown(f"""
        <div class="alert-{box_type}">
            <strong>{icon} {title}</strong><br>
            {content}
        </div>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def create_status_badge(status: str, message: str):
        """Crea un badge de estado"""
        colors = {
            "success": "#28a745",
            "warning": "#ffc107", 
            "error": "#dc3545",
            "info": "#17a2b8"
        }
        
        color = colors.get(status, "#6c757d")
        
        st.markdown(f"""
        <span style="
            background-color: {color}; 
            color: white; 
            padding: 0.25rem 0.75rem; 
            border-radius: 1rem; 
            font-size: 0.875rem;
            display: inline-block;
            margin: 0.25rem 0;
        ">{message}</span>
        """, unsafe_allow_html=True)

# Cargar CSS
load_custom_css()

# CSS personalizado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #1e3c72;
    }
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #1e3c72 0%, #2a5298 100%);
    }
</style>
""", unsafe_allow_html=True)

# Cargar datos usando las nuevas funciones optimizadas
datos_excel, hojas_excel, nombres_hojas = DataLoader.cargar_datos_excel()
datos_word = DataLoader.cargar_documento_word()

# Verificar si los datos se cargaron correctamente
if datos_excel is None or hojas_excel is None or datos_word is None:
    st.error("⚠️ Error al cargar los datos. Verifica que los archivos existan en las rutas especificadas.")
    st.info(f"📁 Buscando archivos en: {ConfigPaths.get_excel_file_path()} y {ConfigPaths.get_word_doc_path()} en Streamlit Cloud")
    st.stop()

# Header principal
st.markdown("""
<div class="main-header">
    <h1>🏛️ REDISEÑO INSTITUCIONAL - MUNICIPIO DE QUIBDÓ</h1>
    <p>Análisis Integral para la Modernización de la Estructura Administrativa Municipal</p>
    <p><strong>Basado en:</strong> Instrumento de Análisis Financiero & Segunda Entrega 7 Septiembre</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/2/21/Flag_of_Colombia.svg/320px-Flag_of_Colombia.svg.png", width=50)
st.sidebar.title("🏛️ REDISEÑO INSTITUCIONAL")
st.sidebar.markdown("**Municipio de Quibdó - Chocó**")
st.sidebar.markdown("---")

seccion = st.sidebar.selectbox(
    "� Selecciona el Análisis:",
    [
        "🏠 Diagnóstico General",
        "📊 Situación Financiera Actual", 
        "🏛️ Estructura Organizacional Vigente",
        "⚖️ Cumplimiento Normativo (Ley 617)",
        "� Indicadores de Gestión Municipal",
        "🎯 Propuesta de Rediseño",
        "� Análisis Costo-Beneficio",
        "📋 Plan de Implementación",
        "📑 Marco Normativo y Legal",
        "📄 Fuentes Documentales"
    ]
)

# Sección: Diagnóstico General
if seccion == "🏠 Diagnóstico General":
    st.header("🏠 Diagnóstico General del Municipio de Quibdó")
    
    # Información contextual enriquecida
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### 📍 Contexto Municipal Detallado
        
        **Municipio:** Quibdó - Capital del Departamento del Chocó  
        **Categoría Municipal:** Según Ley 617 de 2000 (Categoría Especial)  
        **Población aproximada:** 129,237 habitantes (proyección DANE)  
        **Extensión:** 3,337.5 km²  
        **Fundación:** 1648  
        **Altitud:** 53 metros sobre el nivel del mar  
        
        #### 🎯 Marco del Rediseño Institucional
        
        El presente análisis se fundamenta en la **Segunda Entrega del 7 de Septiembre** 
        y el **Instrumento de Análisis Financiero**, documentos que constituyen la base 
        técnica para el rediseño institucional del municipio.
        
        **Objetivo General:** Modernizar la estructura administrativa municipal para:
        
        - ✅ **Optimizar la gestión administrativa** según estándares de eficiencia
        - ✅ **Garantizar el cumplimiento** de la Ley 617 de 2000
        - ✅ **Mejorar la prestación** de servicios a la ciudadanía
        - ✅ **Fortalecer la capacidad** institucional de gestión
        - ✅ **Implementar el MIPG** (Modelo Integrado de Planeación y Gestión)
        - ✅ **Racionalizar los recursos** humanos, técnicos y financieros
        """)
    
    with col2:
        st.info("""
        **📊 Fuentes Primarias de Información:**
        
        � **Instrumento de Análisis Financiero - Grupo 5**  
        • Análisis fiscal y financiero  
        • Indicadores de desempeño  
        • Proyecciones presupuestales  
        
        � **Segunda Entrega 7 Septiembre**  
        • Diagnóstico organizacional  
        • Marco normativo aplicable  
        • Propuestas de mejoramiento  
        
        ⚖️ **Marco Legal de Referencia:**  
        • Ley 617 de 2000  
        • Decreto 1421 de 1993  
        • MIPG - Decreto 1499/2017  
        """)
    
    # Análisis detallado de fuentes
    st.markdown("---")
    st.subheader("📋 Análisis Detallado de las Fuentes de Información")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if datos_excel is not None:
            st.success("✅ **Instrumento de Análisis Financiero - ACTIVO**")
            
            # Análisis específico del Excel
            st.markdown("**📊 Estructura del Instrumento Financiero:**")
            st.write(f"• **Total de hojas de análisis:** {len(nombres_hojas) if nombres_hojas else 0}")
            st.write(f"• **Registros de datos:** {len(datos_excel):,}")
            st.write(f"• **Variables analizadas:** {len(datos_excel.columns)}")
            
            # Análisis de las hojas disponibles
            if nombres_hojas:
                st.markdown("**📑 Módulos del Instrumento:**")
                for i, hoja in enumerate(nombres_hojas[:5], 1):  # Mostrar primeras 5 hojas
                    if hojas_excel and hoja in hojas_excel:
                        registros = len(hojas_excel[hoja])
                        st.write(f"  {i}. **{hoja}** - {registros:,} registros")
                
                if len(nombres_hojas) > 5:
                    st.write(f"  ... y {len(nombres_hojas) - 5} módulos adicionales")
            
            # Análisis de completitud de datos
            datos_completos = datos_excel.count().sum()
            datos_totales = len(datos_excel) * len(datos_excel.columns)
            completitud = (datos_completos / datos_totales * 100) if datos_totales > 0 else 0
            
            st.metric("📈 Completitud de Datos", f"{completitud:.1f}%")
            
            # Identificar columnas financieras clave
            columnas_financieras = [col for col in datos_excel.columns 
                                  if any(term in col.lower() for term in 
                                        ['presupuesto', 'ingreso', 'gasto', 'costo', 'valor', 'monto', 'salario'])]
            
            if columnas_financieras:
                st.markdown("**💰 Variables Financieras Identificadas:**")
                for col in columnas_financieras[:3]:  # Mostrar primeras 3
                    st.write(f"• {col}")
                if len(columnas_financieras) > 3:
                    st.write(f"• ... y {len(columnas_financieras) - 3} variables adicionales")
        else:
            st.error("❌ **Instrumento de Análisis Financiero - NO DISPONIBLE**")
            st.markdown("""
            **⚠️ Impacto en el Análisis:**
            - Limitaciones en el diagnóstico fiscal
            - Imposibilidad de generar indicadores financieros
            - Restricciones en el análisis de sostenibilidad
            """)
    
    with col2:
        if datos_word is not None:
            st.success("✅ **Segunda Entrega 7 Septiembre - ACTIVO**")
            
            # Análisis detallado del documento Word
            st.markdown("**📄 Estructura del Documento Técnico:**")
            st.write(f"• **Total de secciones:** {datos_word.total_parrafos:,}")
            st.write(f"• **Extensión del contenido:** {datos_word.total_palabras:,} palabras")
            st.write(f"• **Tablas de datos:** {datos_word.total_tablas}")
            
            # Análisis de contenido temático
            palabras_clave = datos_word.palabras_clave
            
            # Categorías de mayor relevancia
            temas_principales = {
                'Marco Institucional': palabras_clave.get('institucional', 0) + palabras_clave.get('organización', 0),
                'Aspectos Normativos': palabras_clave.get('ley_617', 0) + palabras_clave.get('normatividad', 0) + palabras_clave.get('decreto', 0),
                'Gestión Administrativa': palabras_clave.get('administrativo', 0) + palabras_clave.get('gestión', 0),
                'Recursos Humanos': palabras_clave.get('personal', 0) + palabras_clave.get('cargo', 0) + palabras_clave.get('funcionario', 0),
                'Aspectos Financieros': palabras_clave.get('financiero', 0) + palabras_clave.get('presupuesto', 0) + palabras_clave.get('fiscal', 0)
            }
            
            st.markdown("**🎯 Énfasis Temático del Documento:**")
            for tema, frecuencia in sorted(temas_principales.items(), key=lambda x: x[1], reverse=True):
                if frecuencia > 0:
                    st.write(f"• **{tema}:** {frecuencia} referencias")
            
            # Densidad de contenido técnico
            total_referencias = sum(palabras_clave.values())
            densidad_tecnica = (total_referencias / datos_word.total_palabras * 100) if datos_word.total_palabras > 0 else 0
            
            st.metric("🔬 Densidad Técnica", f"{densidad_tecnica:.2f}%")
            
            # Palabras clave más frecuentes
            top_palabras = sorted(palabras_clave.items(), key=lambda x: x[1], reverse=True)[:5]
            
            st.markdown("**🏷️ Términos Más Relevantes:**")
            for palabra, freq in top_palabras:
                if freq > 0:
                    st.write(f"• **{palabra.capitalize()}:** {freq}")
        else:
            st.error("❌ **Segunda Entrega 7 Septiembre - NO DISPONIBLE**")
            st.markdown("""
            **⚠️ Impacto en el Análisis:**
            - Ausencia de diagnóstico organizacional
            - Falta de marco conceptual específico
            - Limitaciones en propuestas técnicas
            """)
    
    # Síntesis diagnóstica integrada
    if datos_word is not None:
        st.markdown("---")
        st.subheader("📊 Síntesis Diagnóstica Integrada")
        
        palabras_clave = datos_word.palabras_clave
        
        # Análisis multidimensional
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            valor_normativo = (palabras_clave.get('ley_617', 0) + 
                             palabras_clave.get('normatividad', 0) + 
                             palabras_clave.get('decreto', 0) + 
                             palabras_clave.get('resolución', 0))
            st.metric(
                label="⚖️ Dimensión Normativa",
                value=valor_normativo,
                help="Referencias al marco legal y normativo aplicable"
            )
        
        with col2:
            valor_organizacional = (palabras_clave.get('estructura', 0) + 
                                  palabras_clave.get('organizacional', 0) + 
                                  palabras_clave.get('cargo', 0) + 
                                  palabras_clave.get('dependencia', 0))
            st.metric(
                label="🏛️ Dimensión Organizacional",
                value=valor_organizacional,
                help="Referencias a estructura y organización institucional"
            )
        
        with col3:
            valor_financiero = (palabras_clave.get('financiero', 0) + 
                              palabras_clave.get('presupuesto', 0) + 
                              palabras_clave.get('fiscal', 0) + 
                              palabras_clave.get('recursos', 0))
            st.metric(
                label="💰 Dimensión Financiera",
                value=valor_financiero,
                help="Referencias a aspectos económicos y presupuestales"
            )
        
        with col4:
            valor_gestion = (palabras_clave.get('gestión', 0) + 
                           palabras_clave.get('proceso', 0) + 
                           palabras_clave.get('calidad', 0) + 
                           palabras_clave.get('control', 0))
            st.metric(
                label="📈 Dimensión de Gestión",
                value=valor_gestion,
                help="Referencias a gestión, procesos y control"
            )
        
        # Análisis de prioridades identificadas
        st.subheader("🎯 Prioridades Identificadas en los Documentos")
        
        # Crear visualización de prioridades
        dimensiones = ['Normativa', 'Organizacional', 'Financiera', 'Gestión']
        valores = [valor_normativo, valor_organizacional, valor_financiero, valor_gestion]
        
        fig_dimensiones = px.bar(
            x=dimensiones,
            y=valores,
            title="Distribución de Énfasis por Dimensiones del Rediseño",
            labels={'x': 'Dimensiones de Análisis', 'y': 'Frecuencia de Referencias'},
            color=valores,
            color_continuous_scale='viridis'
        )
        fig_dimensiones.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig_dimensiones, use_container_width=True)
        
        # Hallazgos específicos por fuente
        st.subheader("🔍 Hallazgos Específicos por Fuente")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **📋 Segunda Entrega - Hallazgos Principales:**
            """)
            
            # Identificar las 3 categorías más mencionadas
            categorias_relevantes = {
                'Aspectos Institucionales': palabras_clave.get('institucional', 0) + palabras_clave.get('institución', 0),
                'Marco Municipal': palabras_clave.get('municipio', 0) + palabras_clave.get('municipal', 0),
                'Gestión Administrativa': palabras_clave.get('administrativo', 0) + palabras_clave.get('administración', 0),
                'Desarrollo Organizacional': palabras_clave.get('desarrollo', 0) + palabras_clave.get('modernización', 0),
                'Control y Seguimiento': palabras_clave.get('control', 0) + palabras_clave.get('seguimiento', 0)
            }
            
            for categoria, valor in sorted(categorias_relevantes.items(), key=lambda x: x[1], reverse=True)[:3]:
                if valor > 0:
                    st.write(f"✅ **{categoria}:** {valor} referencias específicas")
        
        with col2:
            if datos_excel is not None:
                st.markdown("""
                **📊 Instrumento Financiero - Capacidades Identificadas:**
                """)
                
                # Análisis de capacidades del Excel
                capacidades = []
                
                if nombres_hojas and len(nombres_hojas) > 0:
                    capacidades.append(f"✅ **Análisis multimensional:** {len(nombres_hojas)} módulos de evaluación")
                
                if len(datos_excel) > 100:
                    capacidades.append(f"✅ **Base de datos robusta:** {len(datos_excel):,} registros para análisis")
                
                # Buscar indicadores financieros
                cols_financieras = [col for col in datos_excel.columns 
                                  if any(term in col.lower() for term in ['valor', 'costo', 'presupuesto'])]
                if cols_financieras:
                    capacidades.append(f"✅ **Indicadores financieros:** {len(cols_financieras)} variables económicas")
                
                # Verificar completitud
                completitud = (datos_excel.count().sum() / (len(datos_excel) * len(datos_excel.columns)) * 100)
                if completitud > 70:
                    capacidades.append(f"✅ **Alta completitud:** {completitud:.1f}% de datos válidos")
                
                for capacidad in capacidades:
                    st.write(capacidad)
            else:
                st.warning("⚠️ **Instrumento Financiero no disponible**")
        
        # Conclusiones del diagnóstico
        st.subheader("� Conclusiones del Diagnóstico General")
        
        conclusiones = []
        
        if valor_normativo > 5:
            conclusiones.append("🎯 **Sólido fundamento normativo** - El documento evidencia conocimiento del marco legal aplicable")
        
        if valor_organizacional > 10:
            conclusiones.append("🏛️ **Enfoque organizacional definido** - Clara orientación hacia la reestructuración institucional")
        
        if datos_excel is not None and nombres_hojas and len(nombres_hojas) > 3:
            conclusiones.append("💰 **Capacidad de análisis financiero** - Herramientas disponibles para evaluación fiscal")
        
        if datos_word.total_palabras > 5000:
            conclusiones.append("📄 **Diagnóstico comprehensivo** - Documento técnico con alcance detallado")
        
        if palabras_clave.get('mipg', 0) > 0:
            conclusiones.append("📊 **Orientación hacia estándares MIPG** - Alineación con mejores prácticas de gestión pública")
        
        # Mostrar conclusiones
        for conclusion in conclusiones:
            st.write(conclusion)
        
        if not conclusiones:
            st.info("💡 **Diagnóstico en construcción** - Se requiere mayor profundidad en las fuentes de información")

# Sección: Situación Financiera Actual
elif seccion == "📊 Situación Financiera Actual":
    st.header("📊 Situación Financiera Actual del Municipio de Quibdó")
    
    if datos_excel is not None:
        st.success("✅ Análisis basado en el Instrumento de Análisis Financiero - Grupo 5")
        
        # Información general del instrumento
        st.markdown("""
        ### 💼 Marco del Análisis Financiero
        
        El **Instrumento de Análisis Financiero** constituye la herramienta técnica fundamental 
        para evaluar la situación fiscal y económica del Municipio de Quibdó, permitiendo:
        
        - 📊 **Análisis de sostenibilidad fiscal** según Ley 617 de 2000
        - 💰 **Evaluación de indicadores financieros** municipales
        - 📈 **Proyección de escenarios** presupuestales
        - ⚖️ **Medición del cumplimiento** de límites normativos
        - 🎯 **Identificación de oportunidades** de optimización
        """)
        
        # Análisis detallado del instrumento
        st.subheader("📑 Estructura Detallada del Instrumento Financiero")
        
        # Información general consolidada
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📄 Módulos de Análisis", len(nombres_hojas) if nombres_hojas else 0)
        with col2:
            st.metric("📊 Total de Registros", f"{len(datos_excel):,}")
        with col3:
            st.metric("📋 Variables Financieras", len(datos_excel.columns))
        with col4:
            # Calcular completitud general
            completitud_general = (datos_excel.count().sum() / (len(datos_excel) * len(datos_excel.columns)) * 100) if len(datos_excel) > 0 else 0
            st.metric("📈 Completitud General", f"{completitud_general:.1f}%")
        
        # Análisis por módulos del instrumento
        if nombres_hojas:
            st.subheader("🔍 Análisis Detallado por Módulo del Instrumento")
            
            # Crear análisis consolidado de todos los módulos
            modulos_info = []
            
            for idx, hoja in enumerate(nombres_hojas):
                if hojas_excel and hoja in hojas_excel:
                    datos_hoja = hojas_excel[hoja]
                    
                    # Análisis específico de cada módulo
                    info_modulo = {
                        'nombre': hoja,
                        'registros': len(datos_hoja),
                        'variables': len(datos_hoja.columns),
                        'completitud': (datos_hoja.count().sum() / (len(datos_hoja) * len(datos_hoja.columns)) * 100) if len(datos_hoja) > 0 else 0,
                        'columnas_numericas': len(datos_hoja.select_dtypes(include=[np.number]).columns),
                        'columnas_financieras': len([col for col in datos_hoja.columns if any(term in col.lower() for term in ['valor', 'costo', 'presupuesto', 'ingreso', 'gasto', 'monto'])])
                    }
                    modulos_info.append(info_modulo)
            
            if modulos_info:
                # Crear DataFrame para visualización
                df_modulos = pd.DataFrame(modulos_info)
                
                # Visualización de la estructura del instrumento
                col1, col2 = st.columns(2)
                
                with col1:
                    # Gráfico de registros por módulo
                    fig_registros = px.bar(
                        df_modulos,
                        x='registros',
                        y='nombre',
                        orientation='h',
                        title='Volumen de Registros por Módulo',
                        labels={'registros': 'Número de Registros', 'nombre': 'Módulo'},
                        color='registros',
                        color_continuous_scale='blues'
                    )
                    fig_registros.update_layout(height=400)
                    st.plotly_chart(fig_registros, use_container_width=True)
                
                with col2:
                    # Gráfico de completitud por módulo
                    fig_completitud = px.bar(
                        df_modulos,
                        x='completitud',
                        y='nombre',
                        orientation='h',
                        title='Completitud de Datos por Módulo (%)',
                        labels={'completitud': 'Completitud (%)', 'nombre': 'Módulo'},
                        color='completitud',
                        color_continuous_scale='greens'
                    )
                    fig_completitud.update_layout(height=400)
                    st.plotly_chart(fig_completitud, use_container_width=True)
                
                # Tabla resumen de módulos
                st.subheader("� Resumen Ejecutivo por Módulo")
                
                # Formatear DataFrame para mostrar
                df_display = df_modulos.copy()
                df_display['completitud'] = df_display['completitud'].round(1).astype(str) + '%'
                df_display.columns = ['Módulo', 'Registros', 'Variables', 'Completitud', 'Var. Numéricas', 'Var. Financieras']
                
                st.dataframe(df_display, use_container_width=True)
            
            # Análisis individual de módulos clave
            st.subheader("🔎 Análisis Individual de Módulos Clave")
            
            for idx, hoja in enumerate(nombres_hojas):
                if hojas_excel and hoja in hojas_excel:
                    datos_hoja = hojas_excel[hoja]
                    
                    with st.expander(f"📊 Módulo: {hoja} ({len(datos_hoja):,} registros)"):
                        
                        # Información básica del módulo
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("📊 Total Registros", f"{len(datos_hoja):,}")
                        with col2:
                            st.metric("📋 Variables", len(datos_hoja.columns))
                        with col3:
                            completitud_modulo = (datos_hoja.count().sum() / (len(datos_hoja) * len(datos_hoja.columns)) * 100) if len(datos_hoja) > 0 else 0
                            st.metric("📈 Completitud", f"{completitud_modulo:.1f}%")
                        
                        # Vista previa de datos
                        st.markdown("**📋 Vista Previa de Datos:**")
                        st.dataframe(datos_hoja.head(3), use_container_width=True)
                        
                        # Análisis de variables financieras específicas
                        columnas_financieras = [col for col in datos_hoja.columns 
                                              if any(term in col.lower() for term in 
                                                    ['valor', 'costo', 'presupuesto', 'ingreso', 'gasto', 'monto', 'salario', 'honorarios'])]
                        
                        if columnas_financieras:
                            st.markdown("**💰 Variables Financieras Identificadas:**")
                            
                            for col_fin in columnas_financieras[:3]:  # Mostrar primeras 3
                                if pd.api.types.is_numeric_dtype(datos_hoja[col_fin]):
                                    valores = datos_hoja[col_fin].dropna()
                                    if len(valores) > 0:
                                        col_val1, col_val2, col_val3, col_val4 = st.columns(4)
                                        with col_val1:
                                            st.write(f"**{col_fin}**")
                                        with col_val2:
                                            st.write(f"Total: ${valores.sum():,.0f}")
                                        with col_val3:
                                            st.write(f"Promedio: ${valores.mean():,.0f}")
                                        with col_val4:
                                            st.write(f"Registros: {len(valores)}")
                            
                            if len(columnas_financieras) > 3:
                                st.write(f"... y {len(columnas_financieras) - 3} variables financieras adicionales")
                        
                        # Análisis de tipos de datos
                        st.markdown("**🔍 Análisis de Tipos de Variables:**")
                        tipos_datos = datos_hoja.dtypes.value_counts()
                        
                        for tipo, cantidad in tipos_datos.items():
                            st.write(f"• **{tipo}:** {cantidad} variables")
        
        # Análisis financiero consolidado
        st.subheader("💰 Análisis Financiero Consolidado Municipal")
        
        st.markdown("""
        ### 📊 Indicadores Financieros Clave
        
        Basado en el procesamiento del Instrumento de Análisis Financiero, 
        se identifican los siguientes componentes del análisis fiscal municipal:
        """)
        
        # Identificar y analizar todas las variables financieras
        todas_variables_financieras = []
        resumen_financiero = {}
        
        if hojas_excel:
            for hoja_nombre, hoja_datos in hojas_excel.items():
                cols_financieras = [col for col in hoja_datos.columns 
                                  if any(term in col.lower() for term in 
                                        ['ingreso', 'gasto', 'presupuesto', 'costo', 'valor', 'monto', 'salario', 'honorarios', 'transferencia'])]
                
                if cols_financieras:
                    st.markdown(f"**📊 Variables Financieras en {hoja_nombre}:**")
                    
                    for col in cols_financieras:
                        todas_variables_financieras.append(f"{hoja_nombre} - {col}")
                        
                        # Análisis estadístico si es numérico
                        if pd.api.types.is_numeric_dtype(hoja_datos[col]):
                            valores = hoja_datos[col].dropna()
                            if len(valores) > 0:
                                resumen_financiero[f"{hoja_nombre}_{col}"] = {
                                    'total': valores.sum(),
                                    'promedio': valores.mean(),
                                    'registros': len(valores),
                                    'categoria': hoja_nombre
                                }
                                
                                st.write(f"  • **{col}:** Total ${valores.sum():,.0f} | Promedio ${valores.mean():,.0f} | {len(valores)} registros")
                        else:
                            st.write(f"  • **{col}:** Variable categórica")
                    
                    st.markdown("---")
        
        # Síntesis del análisis financiero
        if resumen_financiero:
            st.subheader("📈 Síntesis del Análisis Financiero")
            
            # Calcular totales por categoría si es posible
            total_general = sum([info['total'] for info in resumen_financiero.values() if info['total'] > 0])
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("💰 Variables Financieras", len(todas_variables_financieras))
            with col2:
                st.metric("📊 Módulos con Datos $", len(set([info['categoria'] for info in resumen_financiero.values()])))
            with col3:
                if total_general > 0:
                    st.metric("💵 Valor Total Identificado", f"${total_general:,.0f}")
                else:
                    st.metric("💵 Análisis", "Cualitativo")
            
            # Análisis por categorías
            st.markdown("**🎯 Recomendaciones para el Análisis Financiero:**")
            
            recomendaciones = [
                "📊 **Consolidar indicadores** de sostenibilidad fiscal según Ley 617",
                "💰 **Analizar la estructura** de ingresos corrientes vs gastos de funcionamiento",
                "📈 **Evaluar la capacidad** de generación de recursos propios",
                "⚖️ **Verificar cumplimiento** de límites de gastos de personal",
                "🎯 **Identificar oportunidades** de optimización presupuestal",
                "📋 **Proyectar escenarios** de sostenibilidad a mediano plazo"
            ]
            
            for rec in recomendaciones:
                st.write(rec)
        else:
            st.info("💡 Las variables financieras identificadas requieren análisis adicional para generar indicadores específicos")
    
    else:
        st.error("❌ Instrumento de Análisis Financiero no disponible")
        st.markdown("""
        ### ⚠️ Limitaciones del Análisis Sin el Instrumento Financiero
        
        La ausencia del Instrumento de Análisis Financiero impacta significativamente 
        la capacidad de realizar un diagnóstico fiscal completo:
        
        #### 📋 Componentes Faltantes:
        - 📊 **Indicadores de Ley 617:** Límites de gastos de funcionamiento y personal
        - 💰 **Estados financieros:** Ejecución presupuestal e indicadores fiscales
        - 📈 **Análisis de sostenibilidad:** Proyecciones y escenarios fiscales
        - ⚖️ **Cumplimiento normativo:** Verificación de límites legales
        - 🎯 **Benchmarking:** Comparación con otros municipios categoría especial
        
        #### 🔧 Para Completar el Análisis se Requiere:
        - 📄 **Instrumento de Análisis Financiero** (Excel completo)
        - 💰 **Estados financieros municipales** de los últimos 3 años
        - 📊 **Ejecución presupuestal** vigencia actual
        - ⚖️ **Reportes de cumplimiento** Ley 617
        - 📈 **Proyecciones fiscales** a mediano plazo
        - 🏛️ **Información de gastos** por secretarías y dependencias
        """)
        
        # Análisis alternativo basado en el documento Word si está disponible
        if datos_word is not None:
            st.subheader("📄 Referencias Financieras en el Documento Técnico")
            
            palabras_clave = datos_word.palabras_clave
            
            # Métricas financieras del documento
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("💰 Referencias Financieras", palabras_clave.get('financiero', 0))
            with col2:
                st.metric("📊 Referencias Presupuesto", palabras_clave.get('presupuesto', 0))
            with col3:
                st.metric("💵 Referencias Fiscales", palabras_clave.get('fiscal', 0))
            with col4:
                st.metric("📈 Referencias Recursos", palabras_clave.get('recursos', 0))
            
            # Análisis conceptual financiero
            conceptos_financieros = {
                'Presupuesto': palabras_clave.get('presupuesto', 0),
                'Ingresos': palabras_clave.get('ingresos', 0),
                'Gastos': palabras_clave.get('gastos', 0),
                'Inversión': palabras_clave.get('inversión', 0),
                'Recursos': palabras_clave.get('recursos', 0),
                'Costos': palabras_clave.get('costo', 0)
            }
            
            # Filtrar conceptos con valores > 0
            conceptos_activos = {k: v for k, v in conceptos_financieros.items() if v > 0}
            
            if conceptos_activos:
                st.markdown("**📊 Conceptos Financieros Identificados en el Documento:**")
                
                df_conceptos = pd.DataFrame(
                    list(conceptos_activos.items()),
                    columns=['Concepto Financiero', 'Referencias']
                )
                
                fig_conceptos = px.bar(
                    df_conceptos,
                    x='Concepto Financiero',
                    y='Referencias',
                    title='Énfasis Financiero en el Documento Técnico',
                    color='Referencias',
                    color_continuous_scale='blues'
                )
                st.plotly_chart(fig_conceptos, use_container_width=True)
            else:
                st.info("💡 El documento técnico contiene referencias financieras limitadas")

# Sección: Estructura Organizacional Vigente
elif seccion == "🏛️ Estructura Organizacional Vigente":
    st.header("🏛️ Estructura Organizacional Vigente")
    
    st.markdown("""
    ### 📋 Análisis de la Estructura Administrativa Actual
    
    Esta sección analiza la estructura organizacional actual del Municipio de Quibdó 
    identificando fortalezas, debilidades y oportunidades de mejora.
    """)
    
    if datos_word is not None:
        palabras_clave = datos_word.palabras_clave
        
        # Análisis organizacional
        st.subheader("🏢 Componentes Organizacionales Identificados")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("🏛️ Secretarías", palabras_clave.get('secretaria', 0))
        with col2:
            st.metric("👥 Cargos", palabras_clave.get('cargo', 0))
        with col3:
            st.metric("🏢 Dependencias", palabras_clave.get('dependencia', 0))
        with col4:
            st.metric("👨‍💼 Personal", palabras_clave.get('personal', 0))
        
        # Análisis de la planta de personal
        st.subheader("👥 Análisis de la Planta de Personal")
        
        elementos_planta = {
            'Planta de Personal': palabras_clave.get('planta', 0),
            'Clasificación de Cargos': palabras_clave.get('clasificación', 0),
            'Estructura Jerárquica': palabras_clave.get('jerarquía', 0),
            'Competencias': palabras_clave.get('competencia', 0),
            'Funciones': palabras_clave.get('función', 0)
        }
        
        df_planta = pd.DataFrame(
            list(elementos_planta.items()),
            columns=['Aspecto', 'Referencias']
        )
        
        fig_planta = px.bar(
            df_planta,
            x='Referencias',
            y='Aspecto',
            orientation='h',
            title="Elementos de la Planta de Personal Analizados",
            color='Referencias',
            color_continuous_scale='blues'
        )
        st.plotly_chart(fig_planta, use_container_width=True)
        
        # Organigrama conceptual
        st.subheader("📊 Estructura Organizacional Conceptual")
        
        niveles_org = {
            'Nivel Directivo': ['Alcalde', 'Secretarios', 'Directores'],
            'Nivel Asesor': ['Oficina de Planeación', 'Oficina Jurídica', 'Control Interno'],
            'Nivel Operativo': ['Coordinadores', 'Técnicos', 'Auxiliares'],
            'Nivel de Apoyo': ['Administrativos', 'Servicios Generales']
        }
        
        for nivel, cargos in niveles_org.items():
            st.write(f"**{nivel}:**")
            st.write(" • ".join(cargos))
    
    if datos_excel is not None:
        st.subheader("💼 Análisis Cuantitativo de Personal")
        
        # Buscar datos de personal en Excel
        columnas_personal = []
        if hojas_excel:
            for hoja_nombre, hoja_datos in hojas_excel.items():
                cols_personal = [col for col in hoja_datos.columns 
                               if any(term in col.lower() for term in 
                                     ['personal', 'cargo', 'empleado', 'funcionario', 'nomina'])]
                if cols_personal:
                    st.write(f"**📊 Datos de personal en {hoja_nombre}:**")
                    for col in cols_personal:
                        st.write(f"• {col}")

# Sección: Cumplimiento Normativo
elif seccion == "⚖️ Cumplimiento Normativo (Ley 617)":
    st.header("⚖️ Cumplimiento Normativo - Ley 617 de 2000")
    
    st.markdown("""
    ### 📜 Marco Legal de Referencia
    
    **Ley 617 de 2000:** "Por la cual se reforma parcialmente la Ley 136 de 1994, el Decreto Extraordinario 
    1222 de 1986, se adiciona la ley orgánica de presupuesto, el Decreto 1421 de 1993, 
    se dictan otras normas tendientes a fortalecer la descentralización, y se dictan normas 
    para la racionalización del gasto público nacional."
    """)
    
    if datos_word is not None:
        palabras_clave = datos_word.palabras_clave
        
        # Análisis del cumplimiento normativo
        st.subheader("📊 Análisis de Cumplimiento Normativo")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="⚖️ Referencias Ley 617",
                value=palabras_clave.get('ley_617', 0),
                help="Menciones específicas de la Ley 617"
            )
        
        with col2:
            st.metric(
                label="📋 Marco Normativo",
                value=palabras_clave.get('normatividad', 0),
                help="Referencias generales a normatividad"
            )
        
        with col3:
            st.metric(
                label="📄 Decretos/Resoluciones",
                value=palabras_clave.get('decreto', 0) + palabras_clave.get('resolución', 0),
                help="Normativa complementaria"
            )
        
        # Principios de gestión pública
        st.subheader("🎯 Principios de Gestión Pública Identificados")
        
        principios = {
            'Eficiencia': palabras_clave.get('eficiencia', 0),
            'Eficacia': palabras_clave.get('eficacia', 0),
            'Efectividad': palabras_clave.get('efectividad', 0),
            'Transparencia': palabras_clave.get('transparencia', 0),
            'Participación Ciudadana': palabras_clave.get('participación', 0)
        }
        
        df_principios = pd.DataFrame(
            list(principios.items()),
            columns=['Principio', 'Frecuencia']
        )
        
        fig_principios = px.bar(
            df_principios,
            x='Principio',
            y='Frecuencia',
            title="Principios de Gestión Pública en el Análisis",
            color='Frecuencia',
            color_continuous_scale='greens'
        )
        st.plotly_chart(fig_principios, use_container_width=True)
        
        # Requerimientos específicos Ley 617
        st.subheader("📋 Requerimientos Específicos Ley 617")
        
        requerimientos = [
            "🔹 **Límites de gastos de funcionamiento** según categoría municipal",
            "🔹 **Racionalización de la planta de personal** y supresión de cargos",
            "🔹 **Límites a gastos de personal** como porcentaje de ingresos corrientes",
            "🔹 **Restricciones a creación de entidades** descentralizadas",
            "🔹 **Medidas de saneamiento fiscal** para municipios en dificultades",
            "🔹 **Fortalecimiento de la capacidad administrativa** municipal"
        ]
        
        for req in requerimientos:
            st.write(req)

# Sección: Indicadores de Gestión Municipal
elif seccion == "📈 Indicadores de Gestión Municipal":
    st.header("📈 Indicadores de Gestión Municipal")
    
    st.markdown("""
    ### 🎯 Sistema de Seguimiento y Evaluación
    
    Los indicadores de gestión permiten medir el desempeño de la administración municipal 
    y orientar las decisiones hacia el logro de los objetivos institucionales.
    """)
    
    if datos_word is not None:
        palabras_clave = datos_word.palabras_clave
        
        # Componentes del sistema de gestión
        st.subheader("🏗️ Componentes del Sistema de Gestión")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📊 Indicadores", palabras_clave.get('indicador', 0))
        with col2:
            st.metric("🎯 Metas", palabras_clave.get('meta', 0))
        with col3:
            st.metric("🎪 Objetivos", palabras_clave.get('objetivo', 0))
        with col4:
            st.metric("📈 Seguimiento", palabras_clave.get('seguimiento', 0))
        
        # Sistema MIPG
        st.subheader("🏛️ Modelo Integrado de Planeación y Gestión (MIPG)")
        
        componentes_mipg = {
            'MIPG': palabras_clave.get('mipg', 0),
            'Gestión': palabras_clave.get('gestión', 0),
            'Calidad': palabras_clave.get('calidad', 0),
            'Control': palabras_clave.get('control', 0),
            'Evaluación': palabras_clave.get('evaluación', 0)
        }
        
        df_mipg = pd.DataFrame(
            list(componentes_mipg.items()),
            columns=['Componente', 'Frecuencia']
        )
        
        # Crear gráfico polar
        fig_mipg = go.Figure()
        fig_mipg.add_trace(go.Scatterpolar(
            r=list(df_mipg['Frecuencia']),
            theta=list(df_mipg['Componente']),
            fill='toself',
            name='MIPG'
        ))
        fig_mipg.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, max(df_mipg['Frecuencia']) + 2] if df_mipg['Frecuencia'].max() > 0 else [0, 5]
                )),
            showlegend=True,
            title="Componentes MIPG Identificados"
        )
        st.plotly_chart(fig_mipg, use_container_width=True)
        
        # Categorías de indicadores
        st.subheader("📊 Categorías de Indicadores Municipales")
        
        categorias_indicadores = [
            "📈 **Indicadores de Eficiencia:** Relación recursos/resultados",
            "🎯 **Indicadores de Eficacia:** Cumplimiento de metas y objetivos", 
            "⚡ **Indicadores de Efectividad:** Impacto en la ciudadanía",
            "💰 **Indicadores Financieros:** Sostenibilidad fiscal",
            "👥 **Indicadores de Talento Humano:** Gestión del personal",
            "🌍 **Indicadores de Servicio:** Calidad de atención ciudadana"
        ]
        
        for categoria in categorias_indicadores:
            st.write(categoria)

# Sección: Propuesta de Rediseño
elif seccion == "🎯 Propuesta de Rediseño":
    st.header("🎯 Propuesta de Rediseño Institucional")
    
    st.markdown("""
    ### 🔄 Modernización de la Estructura Administrativa
    
    Con base en el diagnóstico realizado, se presenta la propuesta de rediseño institucional 
    que busca optimizar la gestión municipal y mejorar el servicio al ciudadano.
    """)
    
    # Objetivos del rediseño
    st.subheader("🎯 Objetivos del Rediseño")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### 🎯 Objetivos Generales
        
        ✅ **Modernizar** la estructura organizacional  
        ✅ **Optimizar** los procesos administrativos  
        ✅ **Mejorar** la eficiencia operativa  
        ✅ **Fortalecer** el servicio al ciudadano  
        ✅ **Garantizar** el cumplimiento normativo
        """)
    
    with col2:
        st.markdown("""
        #### 🎯 Objetivos Específicos
        
        🔹 Racionalizar la planta de personal  
        🔹 Definir competencias y responsabilidades  
        🔹 Implementar el MIPG  
        🔹 Establecer indicadores de gestión  
        🔹 Mejorar la coordinación institucional
        """)
    
    # Ejes de transformación
    st.subheader("🏗️ Ejes de Transformación")
    
    ejes = [
        {
            "eje": "🏛️ Estructura Organizacional",
            "descripcion": "Redefinición de la estructura con base en procesos misionales",
            "acciones": ["Mapeo de procesos", "Rediseño de dependencias", "Definición de niveles jerárquicos"]
        },
        {
            "eje": "👥 Gestión del Talento Humano", 
            "descripcion": "Optimización de la planta de personal y desarrollo de competencias",
            "acciones": ["Análisis de cargos", "Manual de funciones", "Plan de capacitación"]
        },
        {
            "eje": "📋 Procesos y Procedimientos",
            "descripcion": "Simplificación y digitalización de trámites",
            "acciones": ["Mapeo de procesos", "Eliminación de redundancias", "Automatización"]
        },
        {
            "eje": "📊 Sistema de Gestión",
            "descripcion": "Implementación de herramientas de seguimiento y control",
            "acciones": ["Tablero de indicadores", "Sistema de alertas", "Reportes gerenciales"]
        }
    ]
    
    for eje in ejes:
        with st.expander(f"{eje['eje']} - {eje['descripcion']}"):
            st.write("**Acciones específicas:**")
            for accion in eje['acciones']:
                st.write(f"• {accion}")
    
    if datos_word is not None:
        palabras_clave = datos_word.palabras_clave
        
        # Análisis de elementos de implementación
        st.subheader("📈 Elementos de Implementación Identificados")
        
        elementos_impl = {
            'Diagnóstico': palabras_clave.get('diagnóstico', 0),
            'Análisis': palabras_clave.get('análisis', 0),
            'Reforma': palabras_clave.get('reforma', 0),
            'Modernización': palabras_clave.get('modernización', 0),
            'Estrategia': palabras_clave.get('estrategia', 0)
        }
        
        df_impl = pd.DataFrame(
            list(elementos_impl.items()),
            columns=['Elemento', 'Referencias']
        )
        
        fig_impl = px.bar(
            df_impl,
            x='Elemento',
            y='Referencias',
            title="Elementos de Implementación en el Documento",
            color='Referencias',
            color_continuous_scale='viridis'
        )
        st.plotly_chart(fig_impl, use_container_width=True)

# Sección: Propuesta Técnica
elif seccion == "📋 Propuesta Técnica":
    st.header("📋 Propuesta Técnica de Rediseño")
    
    st.markdown("""
    ### 🔧 Especificaciones Técnicas del Rediseño
    
    La propuesta técnica detalla los aspectos operativos y metodológicos 
    para la implementación del nuevo modelo organizacional.
    """)
    
    # Metodología de implementación
    st.subheader("🛠️ Metodología de Implementación")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### 📋 Fase 1: Diagnóstico y Análisis
        
        ✅ **Evaluación de la situación actual**  
        ✅ **Mapeo de procesos existentes**  
        ✅ **Análisis de brechas normativas**  
        ✅ **Identificación de oportunidades**  
        ✅ **Benchmarking con otros municipios**
        
        **⏱️ Duración:** 4-6 semanas  
        **👥 Responsables:** Equipo técnico + Consultores
        """)
    
    with col2:
        st.markdown("""
        #### 🏗️ Fase 2: Diseño de la Nueva Estructura
        
        ✅ **Definición del nuevo organigrama**  
        ✅ **Elaboración de manuales de funciones**  
        ✅ **Diseño de procesos optimizados**  
        ✅ **Definición de indicadores**  
        ✅ **Plan de capacitación**
        
        **⏱️ Duración:** 6-8 semanas  
        **👥 Responsables:** Equipo de rediseño
        """)
    
    if datos_word is not None:
        st.subheader("📊 Soporte Técnico Identificado")
        
        palabras_clave = datos_word.palabras_clave
        
        elementos_tecnicos = {
            'Metodología': palabras_clave.get('metodología', 0),
            'Técnico': palabras_clave.get('técnico', 0),
            'Implementación': palabras_clave.get('implementación', 0),
            'Desarrollo': palabras_clave.get('desarrollo', 0),
            'Sistema': palabras_clave.get('sistema', 0)
        }
        
        df_tecnico = pd.DataFrame(
            list(elementos_tecnicos.items()),
            columns=['Elemento Técnico', 'Referencias']
        )
        
        fig_tecnico = px.bar(
            df_tecnico,
            x='Elemento Técnico',
            y='Referencias',
            title="Elementos Técnicos Identificados",
            color='Referencias',
            color_continuous_scale='plasma'
        )
        st.plotly_chart(fig_tecnico, use_container_width=True)

# Sección: Cronograma de Actividades
elif seccion == "📅 Cronograma de Actividades":
    st.header("📅 Cronograma de Implementación")
    
    st.markdown("""
    ### ⏰ Plan Temporal de Implementación
    
    Cronograma estructurado para la implementación exitosa del rediseño institucional.
    """)
    
    # Crear cronograma visual simplificado
    fases = [
        {"Fase": "Diagnóstico", "Semanas": 6, "Color": "#1f77b4"},
        {"Fase": "Diseño", "Semanas": 8, "Color": "#ff7f0e"},  
        {"Fase": "Piloto", "Semanas": 4, "Color": "#2ca02c"},
        {"Fase": "Implementación", "Semanas": 8, "Color": "#d62728"},
        {"Fase": "Consolidación", "Semanas": 4, "Color": "#9467bd"}
    ]
    
    df_cronograma = pd.DataFrame(fases)
    
    fig_cronograma = px.bar(
        df_cronograma,
        x='Semanas',
        y='Fase',
        orientation='h',
        title='Cronograma de Implementación (30 semanas)',
        color='Fase',
        text='Semanas'
    )
    
    fig_cronograma.update_traces(texttemplate='%{text} sem', textposition='inside')
    fig_cronograma.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig_cronograma, use_container_width=True)
    
    # Hitos principales
    st.subheader("🎯 Hitos Principales")
    
    hitos = [
        "📋 **Semana 6:** Diagnóstico completado",
        "🏗️ **Semana 14:** Nuevo modelo diseñado",
        "🧪 **Semana 18:** Piloto implementado",
        "🚀 **Semana 26:** Implementación general",
        "✅ **Semana 30:** Proyecto consolidado"
    ]
    
    for hito in hitos:
        st.write(hito)

# Sección: Recursos y Presupuesto
elif seccion == "💰 Recursos y Presupuesto":
    st.header("💰 Recursos y Presupuesto")
    
    st.markdown("""
    ### 💼 Estimación de Recursos Necesarios
    
    Recursos humanos, técnicos y financieros requeridos para el rediseño.
    """)
    
    # Presupuesto estimado
    presupuesto_items = [
        {"Categoría": "👥 Consultoría Externa", "Costo": 180000000, "Porcentaje": 60},
        {"Categoría": "🏛️ Recursos Internos", "Costo": 45000000, "Porcentaje": 15},
        {"Categoría": "📚 Capacitación", "Costo": 30000000, "Porcentaje": 10},
        {"Categoría": "💻 Tecnología", "Costo": 24000000, "Porcentaje": 8},
        {"Categoría": "📋 Documentación", "Costo": 15000000, "Porcentaje": 5},
        {"Categoría": "🎯 Comunicación", "Costo": 6000000, "Porcentaje": 2}
    ]
    
    df_presupuesto = pd.DataFrame(presupuesto_items)
    total_presupuesto = df_presupuesto['Costo'].sum()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("💰 Presupuesto Total", f"${total_presupuesto:,.0f}")
        
        # Tabla de presupuesto
        st.dataframe(
            df_presupuesto.style.format({'Costo': '${:,.0f}'}),
            use_container_width=True
        )
    
    with col2:
        # Gráfico de distribución
        fig_presupuesto = px.pie(
            df_presupuesto,
            values='Costo',
            names='Categoría',
            title="Distribución del Presupuesto"
        )
        st.plotly_chart(fig_presupuesto, use_container_width=True)
    
    # Fuentes de financiación
    st.subheader("🏦 Fuentes de Financiación")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("**🏛️ Recursos Propios (60%)**  \n$180,000,000")
    with col2:
        st.info("**🏢 Cofinanciación (30%)**  \n$90,000,000")
    with col3:
        st.info("**🌐 Cooperación (10%)**  \n$30,000,000")

# Página principal
else:
    st.header("🏛️ Rediseño Institucional - Municipio de Quibdó")
    
    st.markdown("""
    ### 📋 Dashboard de Rediseño Institucional
    
    **Municipio de Quibdó - Departamento del Chocó**
    
    Este dashboard presenta un análisis integral para el rediseño institucional 
    del municipio, basado en documentos técnicos y datos financieros.
    
    #### 📊 Secciones Disponibles:
    
    🏠 **Diagnóstico General** - Contexto y situación actual  
    📊 **Situación Financiera** - Análisis económico municipal  
    🏛️ **Estructura Organizacional** - Análisis de la estructura vigente  
    ⚖️ **Cumplimiento Normativo** - Marco legal Ley 617  
    📈 **Indicadores de Gestión** - Métricas y seguimiento  
    🎯 **Propuesta de Rediseño** - Nuevo modelo organizacional  
    📋 **Propuesta Técnica** - Especificaciones de implementación  
    📅 **Cronograma** - Plan temporal de actividades  
    💰 **Recursos y Presupuesto** - Estimación de costos y financiación  
    
    #### 🎯 Objetivos del Rediseño:
    
    ✅ Modernizar la estructura administrativa  
    ✅ Optimizar procesos y procedimientos  
    ✅ Mejorar la eficiencia operativa  
    ✅ Fortalecer el servicio al ciudadano  
    ✅ Garantizar el cumplimiento normativo
    """)
    
    # Estado de los documentos
    col1, col2 = st.columns(2)
    
    with col1:
        if datos_excel is not None:
            st.success("✅ **Instrumento de Análisis Financiero**  \nDocumento Excel cargado correctamente")
        else:
            st.error("❌ **Instrumento de Análisis Financiero**  \nDocumento no disponible")
    
    with col2:
        if datos_word is not None:
            st.success("✅ **Segunda Entrega - Documento Técnico**  \nDocumento Word cargado correctamente")
        else:
            st.error("❌ **Segunda Entrega - Documento Técnico**  \nDocumento no disponible")
    
    if datos_excel is not None or datos_word is not None:
        st.info("💡 **Los documentos están cargados. Selecciona una sección del menú lateral para comenzar el análisis.**")
    else:
        st.warning("⚠️ **Algunos documentos no están disponibles. El análisis será limitado.**")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    📊 Dashboard de Análisis - Municipio de Quibdó | 
    📅 Septiembre 2025 | 
    🔧 Desarrollado con Streamlit
</div>
""", unsafe_allow_html=True)
