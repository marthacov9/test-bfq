from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
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
    genero = st.selectbox("Género", ["Masculino", "Femenino", "Otro"])
with col2:
    edad = st.number_input("Edad", min_value=10, max_value=100, value=25)
    municipio = st.text_input("Municipio / Ciudad")
    fecha = st.date_input("Fecha")

# Campo obligatorio para el correo de contacto del usuario
correo_destino = st.text_input(
    "Correo electrónico de contacto"
)

st.markdown("---")
st.subheader("Instrucciones")
st.write(
    "Califica cada afirmación de **1 (Completamente falso para mí)** a **5 (Completamente verdadero para mí)**."
)


# Cargar las preguntas desde la hoja "Preguntas"
@st.cache_data
def cargar_preguntas():
    df = pd.read_excel("BFQ.xlsx", sheet_name="Preguntas")
    preguntas = []
    for i in range(len(df)):
        val_col = str(df.iloc[i, 2])  # Columna C (índice 2)
        if val_col.startswith("P") and any(char.isdigit() for char in val_col):
            texto_pregunta = str(df.iloc[i, 3])  # Columna D (índice 3)
            preguntas.append({"id": val_col, "texto": texto_pregunta, "row": i})
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

        submitted = st.form_submit_button("Enviar y Procesar Resultados")

    if submitted:
        if not nombre or not correo_destino:
            st.warning(
                "Por favor ingresa tu nombre y tu correo electrónico antes de enviar."
            )
        else:
            try:
                wb = openpyxl.load_workbook("BFQ.xlsx")

                # Función auxiliar robusta para escribir en celdas normales o combinadas (MergedCell)
                def escribir_celda(ws, ref, valor):
                    cell = ws[ref]
                    if type(cell).__name__ == "MergedCell":
                        for merged_range in ws.merged_cells.ranges:
                            if ref in merged_range:
                                min_col, min_row, _, _ = merged_range.bounds
                                ws.cell(row=min_row, column=min_col).value = valor
                                return
                    else:
                        cell.value = valor

                # --- 1. VACIADO EN HOJA "Preguntas" ---
                ws_preg = wb["Preguntas"]
                escribir_celda(ws_preg, "E5", nombre)
                escribir_celda(ws_preg, "E6", ocupacion)
                escribir_celda(ws_preg, "E7", str(fecha))
                escribir_celda(ws_preg, "H7", municipio)
                escribir_celda(ws_preg, "H8", edad)

                # Mapeo y asignación correcta del género en la hoja Preguntas (N6 para fórmulas y K6/K7 para texto visual)
                if genero == "Masculino":
                    val_genero_idx = 2
                    escribir_celda(ws_preg, "K6", "Masculino")
                elif genero == "Femenino":
                    val_genero_idx = 3
                    escribir_celda(ws_preg, "K6", "Femenino")
                else:
                    val_genero_idx = 1
                    escribir_celda(ws_preg, "K6", "Otro")

                escribir_celda(ws_preg, "N6", val_genero_idx)
                escribir_celda(ws_preg, "K7", genero)

                # --- 2. VACIADO EN HOJA "PD con plantilla" ---
                if "PD con plantilla" in wb.sheetnames:
                    ws_pd = wb["PD con plantilla"]
                    escribir_celda(ws_pd, "C6", nombre)
                    escribir_celda(ws_pd, "C7", ocupacion)
                    escribir_celda(ws_pd, "C8", str(fecha))
                    escribir_celda(ws_pd, "E8", municipio)
                    escribir_celda(ws_pd, "C9", genero)
                    escribir_celda(ws_pd, "E9", edad)

                # Inyectar respuestas exclusivamente en las filas de preguntas
                for p in lista_preguntas:
                    row_idx = p["row"] + 1
                    val_elegido = st.session_state[p["id"]]
                    fila_numeros = row_idx + 4

                    # Limpiar opciones anteriores en la fila de números (columnas D a H)
                    for col_c in range(4, 9):
                        cell = ws_preg.cell(row=fila_numeros, column=col_c)
                        if type(cell).__name__ != "MergedCell":
                            cell.value = None

                    # Colocar la "X" visual en la opción elegida
                    col_map = {5: 4, 4: 5, 3: 6, 2: 7, 1: 8}
                    target_col = col_map.get(val_elegido)
                    if target_col:
                        target_cell = ws_preg.cell(row=fila_numeros, column=target_col)
                        if type(target_cell).__name__ != "MergedCell":
                            target_cell.value = "X"

                    # Alimentar el valor numérico en la columna O para activar las fórmulas de subdimensiones y gráficas
                    cell_o = ws_preg.cell(row=fila_numeros, column=15)
                    if type(cell_o).__name__ != "MergedCell":
                        cell_o.value = val_elegido

                # Guardar el archivo procesado temporalmente
                output_path = f"Informe_BFQ_{nombre.replace(' ', '_')}.xlsx"
                wb.save(output_path)

                # --- CONFIGURACIÓN DE ENVÍO EXCLUSIVO AL ADMINISTRADOR ---
                remitente = st.secrets["email"]["remitente"]
                password = st.secrets["email"]["password_app"]
                correo_receptor = st.secrets["email"]["correo_destino"]  # Tu correo configurado

                msg = MIMEMultipart()
                msg["From"] = remitente
                msg["To"] = correo_receptor  # Único destinatario
                msg["Subject"] = f"Nuevo Resultado Cuestionario Big Five - {nombre}"

                cuerpo = (
                    f"Hola,\n\nSe ha completado un nuevo cuestionario de Big Five.\n\n"
                    f"Datos del Evaluado:\n"
                    f"- Nombre: {nombre}\n"
                    f"- Ocupación: {ocupacion}\n"
                    f"- Edad: {edad}\n"
                    f"- Municipio: {municipio}\n"
                    f"- Género: {genero}\n"
                    f"- Correo de contacto del usuario: {correo_destino}\n\n"
                    f"Adjunto encontrarás el archivo de Excel completo con los datos del candidato, cálculos de dimensiones y subdimensiones según el género, y la hoja de gráfica generada automáticamente."
                )
                msg.attach(MIMEText(cuerpo, "plain"))

                with open(output_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {output_path}",
                )
                msg.attach(part)

                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(remitente, password)
                server.sendmail(remitente, correo_receptor, msg.as_string())
                server.quit()

                st.success(
                    "¡Cuestionario enviado con éxito! Sus respuestas han sido procesadas y los resultados se han remitido de forma privada a su bandeja."
                )
                st.balloons()

            except Exception as e:
                st.error(
                    f"Ocurrió un error al procesar o enviar el correo: {e}"
                )
