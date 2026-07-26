from datetime import datetime
import os
import openpyxl
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Tesorería de Curso", page_icon="💰", layout="centered"
)

# Estilo CSS avanzado: Forzando botones rosados brillantes y tipografía legible
st.markdown(
    """
    <style>
    /* Tipografía general limpia y legible para mayor accesibilidad */
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
        font-size: 16px !important;
    }
    
    /* Fondo general de la aplicación */
    .stApp {
        background-color: #fff0f3;
    }
    
    /* Fondo de la barra lateral */
    [data-testid="stSidebar"] {
        background-color: #ffd1dc;
    }
    
    /* Textos generales y títulos en tonos rosados oscuros/vinotinto */
    h1, h2, h3, p, span, label, div {
        color: #5c2c3e !important;
    }
    
    /* CORRECCIÓN BOTONES: Forzar fondo rosa fuerte y texto blanco sin excepciones */
    div.stButton > button, .stButton > button, button[kind="primary"], div.stFormSubmitButton > button {
        width: 100% !important;
        background-color: #ff4d6d !important;
        background-image: none !important;
        color: #ffffff !important;
        font-weight: bold !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        padding: 0.6rem !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(255, 77, 109, 0.3) !important;
    }
    
    div.stButton > button * , .stButton > button * , div.stFormSubmitButton > button * {
        color: #ffffff !important;
    }

    div.stButton > button:hover, .stButton > button:hover, div.stFormSubmitButton > button:hover {
        background-color: #ff1f4b !important;
        color: #ffffff !important;
    }
    
    /* Tarjeta de métricas */
    [data-testid="stMetric"] {
        background-color: #ffccd5;
        padding: 15px;
        border-radius: 12px;
        border: 1px solid #ffb3c1;
    }
    [data-testid="stMetricValue"], [data-testid="stMetricLabel"] {
        color: #5c2c3e !important;
    }
    
    /* Cajas de texto y selectores más grandes y legibles */
    .stTextInput input, .stSelectbox select, .stDateInput input {
        background-color: #ffffff !important;
        color: #5c2c3e !important;
        font-size: 16px !important;
        border-radius: 8px !important;
        border: 1px solid #ffb3c1 !important;
    }
    
    /* Radio buttons del menú lateral y filtros */
    .stRadio label {
        color: #5c2c3e !important;
        font-weight: 600 !important;
        font-size: 16px !important;
    }
    
    /* Botón de descarga de Excel también rosado y armónico */
    .stDownloadButton>button {
        background-color: #ff758c !important;
        color: #ffffff !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        font-weight: bold !important;
        border: none !important;
    }
    .stDownloadButton>button * {
        color: #ffffff !important;
    }
    .stDownloadButton>button:hover {
        background-color: #ff4d6d !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Archivo de Excel donde se guardarán los datos
EXCEL_FILE = "tesoreria_curso.xlsx"

# Lista oficial de alumnos del curso
LISTA_ALUMNOS = [
    "Maximiliano Neira",
    "Victoria Cifuentes",
    "Antonella Enrriquez",
    "Sofia Reyes",
    "Cristobal Canales",
    "Cristobal Saavedra",
    "Cristobal Cortes",
    "Benjamin Candia",
    "Lied Chamber",
    "Juaquin Uriol",
    "Sofia Duran",
    "Fabiana Quijada",
    "Julieta Cordero",
    "Olivia Verdugo",
    "Diego Araujo",
    "Diego Rivera",
    "Dafne Valdes",
    "Emilia Silva",
    "Gianella Hidalgo",
    "Agustin Rebolledo",
    "Agustin Guerra",
    "Maite Guerrero",
    "Samantha Cedeño",
    "Christian Morales",
    "Angel Cuevas",
    "Martina Tobar",
    "Emily Dias",
    "Belen Guitiere",
    "Leon Burgos",
    "Jose Lizondo",
    "Jose Mariman",
    "Bastian Yefi",
    "Vicente Tapia",
    "Rocío Valenzuela",
    "Carlos Alvarez",
    "Abril Alvarez",
    "Mateo Matus",
    "Bruno Cornejo",
    "Gabriel Galaz",
    "Bastian Navarro",
]

# Diccionario para traducir meses a español de forma limpia
MESES_ESPANOL = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def cargar_datos():
  if os.path.exists(EXCEL_FILE):
    try:
      df = pd.read_excel(EXCEL_FILE)
      columnas_necesarias = [
          "Fecha",
          "Nombre Alumno/a",
          "Medio de Pago",
          "Monto ($)",
          "Motivo",
      ]
      for col in columnas_necesarias:
        if col not in df.columns:
          df[col] = ""
      return df[columnas_necesarias]
    except Exception:
      pass

  df = pd.DataFrame(
      columns=["Fecha", "Nombre Alumno/a", "Medio de Pago", "Monto ($)", "Motivo"]
  )
  df.to_excel(EXCEL_FILE, index=False)
  return df


# Cargar datos actuales
df_registros = cargar_datos()

# --- MENÚ LATERAL DE NAVEGACIÓN ---
st.sidebar.title("Navegación")
menu = st.sidebar.radio(
    "Selecciona una opción:", ["Registrar Pago", "Visualizar Estado"]
)

# ==========================================
# OPCIÓN 1: AGREGAR / REGISTRAR PAGO
# ==========================================
if menu == "Registrar Pago":
  st.title("Control de Tesorería Escolar")
  st.markdown("Ingresa los datos del pago o aporte recibido.")

  with st.form("form_pago", clear_on_submit=True):
    fecha = st.date_input("Fecha de Pago", value=datetime.today())

    # Selector desplegable con la lista de alumnos
    nombre = st.selectbox(
        "Nombre del Alumno/a",
        options=sorted(LISTA_ALUMNOS),
        index=0,
    )

    medio_pago = st.selectbox(
        "Medio de Pago", ["Efectivo", "Transferencia", "Otro"]
    )
    monto = st.number_input("Monto ($)", min_value=0, step=500, value=5000)
    motivo = st.text_input(
        "Motivo", placeholder="Ej: Cuota mensual marzo, rifa, paseo..."
    )

    submitted = st.form_submit_button("Guardar Registro")

    if submitted:
      if not motivo.strip():
        st.error("Por favor completa el motivo del pago.")
      else:
        nuevo_registro = pd.DataFrame({
            "Fecha": [str(fecha)],
            "Nombre Alumno/a": [nombre],
            "Medio de Pago": [medio_pago],
            "Monto ($)": [monto],
            "Motivo": [motivo.strip()],
        })

        df_registros = pd.concat(
            [df_registros, nuevo_registro], ignore_index=True
        )
        df_registros.to_excel(EXCEL_FILE, index=False)

        st.success(f"Se guardó correctamente el registro de **{nombre}**.")

# ==========================================
# OPCIÓN 2: VISUALIZAR ESTADO Y ELIMINAR
# ==========================================
elif menu == "Visualizar Estado":
  st.title("Estado de Tesorería")
  st.markdown("Revisa el resumen general por periodos o filtra por alumno.")

  if df_registros.empty:
    st.info("Aún no hay registros guardados en el sistema.")
  else:
    # Procesar fechas para filtros de mes y año
    df_registros["Fecha_dt"] = pd.to_datetime(
        df_registros["Fecha"], errors="coerce"
    )
    df_registros["Año"] = df_registros["Fecha_dt"].dt.year
    df_registros["Mes_Num"] = df_registros["Fecha_dt"].dt.month
    df_registros["Mes"] = df_registros["Mes_Num"].map(MESES_ESPANOL)

    # Sub-opción para ver general, por persona o eliminar registros
    tipo_vista = st.radio(
        "Tipo de Visualización:",
        ["General (Por Meses)", "Por Persona", "Eliminar Registro"],
        horizontal=True,
    )

    if tipo_vista == "General (Por Meses)":
      st.subheader("Historial General y Filtro por Mes")

      anos_disponibles = sorted(
          df_registros["Año"].dropna().unique().astype(int).tolist(),
          reverse=True,
      )

      if anos_disponibles:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
          ano_seleccionado = st.selectbox("Año:", anos_disponibles)

        meses_en_ano = (
            df_registros[df_registros["Año"] == ano_seleccionado]["Mes_Num"]
            .dropna()
            .unique()
            .tolist()
        )
        nombres_meses = ["Todos"] + [
            MESES_ESPANOL[m] for m in sorted(meses_en_ano) if m in MESES_ESPANOL
        ]

        with col_f2:
          mes_seleccionado = st.selectbox("Mes:", nombres_meses)

        df_filtrado = df_registros[df_registros["Año"] == ano_seleccionado]
        if mes_seleccionado != "Todos":
          df_filtrado = df_filtrado[df_filtrado["Mes"] == mes_seleccionado]

        total_periodo = df_filtrado["Monto ($)"].sum()
        etiqueta_metrica = (
            f"Total Recaudado ({mes_seleccionado} {ano_seleccionado})"
            if mes_seleccionado != "Todos"
            else f"Total Recaudado Año {ano_seleccionado}"
        )
        st.metric(etiqueta_metrica, f"${total_periodo:,.0f}")
        st.divider()

        columnas_mostrar = [
            "Fecha",
            "Nombre Alumno/a",
            "Medio de Pago",
            "Monto ($)",
            "Motivo",
        ]
        st.dataframe(
            df_filtrado[columnas_mostrar].reset_index(drop=True),
            use_container_width=True,
        )
      else:
        st.dataframe(
            df_registros[
                [
                    "Fecha",
                    "Nombre Alumno/a",
                    "Medio de Pago",
                    "Monto ($)",
                    "Motivo",
                ]
            ],
            use_container_width=True,
        )

    elif tipo_vista == "Por Persona":
      st.subheader("Búsqueda por Alumno/a")
      alumno_seleccionado = st.selectbox(
          "Selecciona al alumno/a:", sorted(LISTA_ALUMNOS)
      )

      df_persona = df_registros[
          df_registros["Nombre Alumno/a"] == alumno_seleccionado
      ]
      total_persona = df_persona["Monto ($)"].sum()

      st.markdown(
          f"**Aportes totales de {alumno_seleccionado}:**"
          f" **${total_persona:,.0f}**"
      )
      columnas_mostrar = [
          "Fecha",
          "Nombre Alumno/a",
          "Medio de Pago",
          "Monto ($)",
          "Motivo",
      ]
      st.dataframe(
          df_persona[columnas_mostrar].reset_index(drop=True),
          use_container_width=True,
      )

    else:
      st.subheader("Gestión de Eliminación de Registros")
      st.markdown(
          "Selecciona el registro exacto que deseas eliminar (por ejemplo,"
          " pruebas o errores)."
      )

      df_registros["Identificador_Fila"] = (
          df_registros.index.astype(str)
          + " - "
          + df_registros["Fecha"].astype(str)
          + " | "
          + df_registros["Nombre Alumno/a"].astype(str)
          + " | $"
          + df_registros["Monto ($)"].astype(str)
          + " ("
          + df_registros["Motivo"].astype(str)
          + ")"
      )

      fila_a_borrar = st.selectbox(
          "Selecciona el registro a borrar:",
          options=df_registros["Identificador_Fila"].tolist(),
      )

      if st.button("Eliminar este registro"):
        idx_original = int(fila_a_borrar.split(" - ")[0])
        df_registros = df_registros.drop(idx_original)

        columnas_guardar = [
            "Fecha",
            "Nombre Alumno/a",
            "Medio de Pago",
            "Monto ($)",
            "Motivo",
        ]
        df_registros[columnas_guardar].to_excel(EXCEL_FILE, index=False)

        st.success("¡Registro eliminado con éxito! Actualiza la vista para ver el cambio.")
        st.rerun()

    st.divider()

    # Botón para descargar el excel actualizado
    with open(EXCEL_FILE, "rb") as file:
      st.download_button(
          label="Descargar Planilla Excel Actualizada",
          data=file,
          file_name="tesoreria_curso.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheet.sheet",
      )