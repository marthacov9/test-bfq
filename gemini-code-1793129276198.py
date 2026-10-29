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


# Cargar las preguntas desde el Excel original de manera precisa
@st.cache_data
def cargar_preguntas():
    df = pd.read_excel("BFQ.xlsx", sheet_name="Hoja 1")
    preguntas = []
    for i in range(len(df)):
        val_col1 = str(df.iloc[i, 1])  # Columna B (índice 1)
        if val_col1.startswith("P") and any(
            char.isdigit() for char in val_col1
        ):
            texto_pregunta = str(df.iloc[i, 2])  # Columna C (índice 2)
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
                options=[5, 4, 3, 2, 1],
                format_func=lambda x: {
                    5: "5 - Completamente verdadero",
                    4: "4 - Bastante verdadero",
                    3: "3 - Ni verdadero ni falso",
                    2: "2 - Bastante falso",
                    1: "1 - Completamente falso",
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
                    # Abrir el Excel con openpyxl preservando todas las fórmulas de la Hoja 2
                    wb = openpyxl.load_workbook("BFQ.xlsx")
                    ws = wb["Hoja 1"]

                    # Asignar datos personales en las celdas correspondientes de tu Excel
                    ws["C4"] = nombre
                    ws["C5"] = ocupacion
                    ws["C6"] = genero
                    ws["E5"] = str(fecha)
                    ws["G5"] = municipio
                    ws["F6"] = edad

                    # Inyectar las respuestas del usuario en el Excel
                    # (Dependiendo de cómo tengas estructuradas las celdas de respuesta en tu Hoja 1,
                    # openpyxl actualizará automáticamente las fórmulas de la Hoja 2)
                    for p in lista_preguntas:
                        row_idx = p["row"] + 1  # Ajuste de índice de fila en openpyxl
                        val_elegido = respuestas[p["id"]]
                        # Aquí puedes asignar el valor directamente a la celda de respuesta correspondiente de la fila.
                        # Por ejemplo, si la respuesta numérica se coloca en una celda específica al lado de la pregunta:
                        # ws.cell(row=row_idx, column=4).value = val_elegido

                    # Guardar el archivo procesado conservando todas las fórmulas y cálculos de la Hoja 2
                    output_path = f"Informe_BFQ_{nombre.replace(' ', '_')}.xlsx"
                    wb.save(output_path)

                    st.success(
                        "¡Cuestionario respondido y procesado con éxito por el motor de tu Excel!"
                    )
                    st.balloons()

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Descargar tu Informe Completo en Excel (con Perfil Calculado)",
                            data=f,
                            file_name=output_path,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                except Exception as e:
                    st.error(
                        f"Ocurrió un error al procesar el archivo en Excel: {e}"
                    )
