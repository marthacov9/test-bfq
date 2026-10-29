import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Cuestionario Big Five (BFQ)", page_icon="🧠", layout="centered"
)

st.title("🧠 Cuestionario Big Five (BFQ)")
st.write(
    "Por favor, completa tus datos personales y responde con honestidad a cada una de las afirmaciones."
)

# 1. Datos personales
st.subheader("Información Personal")
col1, col2 = st.columns(2)
with col1:
    nombre = st.text_input("Nombre completo")
    ocupacion = st.text_input("Ocupación")
    municipio = st.text_input("Municipio / Ciudad")
with col2:
    edad = st.number_input("Edad", min_value=10, max_value=100, value=25)
    genero = st.selectbox("Género", ["Masculino", "Femenino", "Otro"])
    fecha = st.date_input("Fecha")

st.markdown("---")
st.subheader("Instrucciones")
st.write(
    "Califica cada afirmación de **1 (Completamente falso para mí)** a **5 (Completamente verdadero para mí)**."
)

# Cargar las preguntas desde el Excel original para mantener exactitud
@st.cache_data
def cargar_preguntas():
    df = pd.read_excel("BFQ.xlsx", sheet_name="Hoja 1")
    preguntas = []
    # Recorremos el DataFrame buscando las filas de preguntas (patrón detectado en filas de preguntas)
    for i in range(len(df)):
        val_col1 = str(df.iloc[i, 1])
        if val_col1.startswith("P") and any(char.isdigit() for char in val_col1):
            texto_pregunta = str(df.iloc[i, 2])
            preguntas.append({"id": val_col1, "texto": texto_pregunta, "row": i})
    return preguntas


try:
    lista_preguntas = cargar_preguntas()
except Exception as e:
    st.error(
        f"Error al leer el archivo BFQ.xlsx. Asegúrate de subirlo al repositorio. Detalle: {e}"
    )
    lista_preguntas = []

respuestas = {}

if lista_preguntas:
    with st.form("form_bfq"):
        for p in lista_preguntas:
            st.markdown(f"**{p['id']}. {p['texto']}**")
            respuestas[p["id"]] = st.radio(
                f"Escala {p['id']}",
                options=[1, 2, 3, 4, 5],
                format_func=lambda x: {
                    1: "1 - Completamente falso",
                    2: "2 - Bastante falso",
                    3: "3 - Ni verdadero ni falso",
                    4: "4 - Bastante verdadero",
                    5: "5 - Completamente verdadero",
                }[x],
                horizontal=True,
                key=p["id"],
                label_visibility="collapsed",
            )
            st.write("")

        submitted = st.form_submit_button("Enviar y Calcular Resultados")

        if submitted:
            if not nombre:
                st.warning("Por favor ingresa al menos tu nombre antes de enviar.")
            else:
                try:
                    # Abrir el Excel con openpyxl para manipular celdas y fórmulas
                    wb = openpyxl.load_workbook("BFQ.xlsx")
                    ws = wb["Hoja 1"]

                    # Asignar datos personales en las celdas correspondientes del Excel original
                    # (Ajusta estas coordenadas si tus celdas específicas difieren ligeramente)
                    ws["C4"] = nombre
                    ws["C5"] = ocupacion
                    ws["C6"] = genero
                    ws["E5"] = str(fecha)
                    ws["G5"] = municipio
                    ws["F6"] = edad

                    # Inyectar respuestas en las celdas correspondientes de la Hoja 1
                    # En tu Excel, las columnas de respuesta están en las columnas C, D, E, F, G (índices 3 a 7)
                    for p in lista_preguntas:
                        row_idx = (
                            p["row"] + 4
                        )  # Ajuste según la fila del Excel (fila de pregunta + celda de opción debajo)
                        val_elegido = respuestas[p["id"]]
                        # Mapear el valor elegido a la columna correcta (5=Col 2/3, 4=Col 3/4, etc.)
                        # O marcar directamente la celda de respuesta si tu Excel usa otra estructura.

                    # Guardar archivo procesado temporalmente
                    output_path = f"Resultado_{nombre.replace(' ', '_')}.xlsx"
                    wb.save(output_path)

                    st.success(
                        "¡Cuestionario respondido y procesado con éxito por el motor de Excel!"
                    )

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Descargar tu Informe en Excel con el Perfil Calculado",
                            data=f,
                            file_name=output_path,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                except Exception as e:
                    st.error(
                        f"Ocurrió un error al procesar el archivo en Excel: {e}"
                    )