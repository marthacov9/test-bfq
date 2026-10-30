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

# Campo obligatorio para el correo de contacto o el tuyo institucional
correo_destino = st.text_input(
    "Correo electrónico (para recibir tus resultados o de contacto)"
)

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

        submitted = st.form_submit_button("Enviar y Procesar Resultados")

    if submitted:
        if not nombre or not correo_destino:
            st.warning(
                "Por favor ingresa tu nombre y tu correo electrónico antes de enviar."
            )
        else:
            try:
                # Abrir el Excel con openpyxl preservando todas las fórmulas de la Hoja 2
                wb = openpyxl.load_workbook("BFQ.xlsx")
                ws = wb["Hoja 1"]

                # Asignar datos personales en las celdas exactas de tu Excel
                ws["B5"] = nombre
                ws["B6"] = ocupacion
                ws["D7"] = str(fecha)
                ws["F7"] = municipio
                ws["B8"] = genero
                ws["F8"] = edad

                # Inyectar las respuestas del usuario en el Excel
                for p in lista_preguntas:
                    row_idx = p["row"] + 1
                    val_elegido = st.session_state[p["id"]]

                    for col_c in range(3, 8):
                        ws.cell(row=row_idx + 4, column=col_c).value = None

                    col_map = {5: 3, 4: 4, 3: 5, 2: 6, 1: 7}
                    target_col = col_map.get(val_elegido)
                    if target_col:
                        ws.cell(row=row_idx + 4, column=target_col).value = "X"

                # Guardar temporalmente el archivo procesado en el servidor
                output_path = f"Informe_BFQ_{nombre.replace(' ', '_')}.xlsx"
                wb.save(output_path)

                # --- CONFIGURACIÓN DE ENVÍO POR CORREO ---
                remitente = st.secrets["email"]["remitente"]
                password = st.secrets["email"]["password_app"]
                
                # CORREO DESTINO: 
                # Apunta a tu correo fijo de administración configurado en los secrets
                correo_receptor = st.secrets["email"]["correo_destino"]

                msg = MIMEMultipart()
                msg["From"] = remitente
                msg["To"] = correo_receptor
                msg["Subject"] = f"Nuevo Resultado Cuestionario Big Five - {nombre}"

                cuerpo = (
                    f"Hola,\n\nSe ha completado un nuevo cuestionario de Big Five.\n\n"
                    f"Datos del Evaluado:\n"
                    f"- Nombre: {nombre}\n"
                    f"- Ocupación: {ocupacion}\n"
                    f"- Edad: {edad}\n"
                    f"- Municipio: {municipio}\n"
                    f"- Correo de contacto: {correo_destino}\n\n"
                    f"Adjunto encontrarás el archivo de Excel con el perfil calculado."
                )
                msg.attach(MIMEText(cuerpo, "plain"))

                # Adjuntar el archivo Excel (Error de sintaxis corregido correctamente)
                with open(output_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {output_path}",
                )
                msg.attach(part)

                # Conectar al servidor SMTP de Gmail y enviar
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(remitente, password)
                server.sendmail(remitente, correo_receptor, msg.as_string())
                server.quit()

                st.success(
                    "¡Cuestionario enviado con éxito! Los resultados han sido procesados y enviados por correo."
                )
                st.balloons()

            except Exception as e:
                st.error(
                    f"Ocurrió un error al procesar o enviar el correo (verifica tus credenciales de correo): {e}"
                )
