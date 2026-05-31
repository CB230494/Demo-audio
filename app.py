import streamlit as st
import tempfile
import os
import re
import io
from datetime import datetime
import whisper
from streamlit_mic_recorder import mic_recorder
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# =========================
# CONFIGURACIÓN
# =========================

st.set_page_config(
    page_title="Transcriptor de Reuniones",
    page_icon="🎙️",
    layout="wide"
)

LIMITE_MINUTOS = 10
LIMITE_SEGUNDOS = LIMITE_MINUTOS * 60

# =========================
# MODELO WHISPER
# =========================

@st.cache_resource
def cargar_modelo():
    return whisper.load_model("base")

modelo = cargar_modelo()

# =========================
# SESSION STATE
# =========================

if "transcripcion_total" not in st.session_state:
    st.session_state.transcripcion_total = ""

if "texto_corregido" not in st.session_state:
    st.session_state.texto_corregido = ""

# =========================
# FUNCIONES
# =========================

def corregir_texto_basico(texto):
    if not texto:
        return ""

    correcciones = {
        "genecis": "Génesis",
        "genesis": "Génesis",
        "asocion": "asociación",
        "reunion": "reunión",
        "reuniones": "reuniones",
        "contaminacion": "contaminación",
        "dia": "día",
        "veintiseis": "veintiséis",
        "atencion": "atención",
        "revision": "revisión",
        "senora": "señora",
        "direccion": "dirección",
        "desarrollo": "desarrollo",
        "comite": "comité",
        "sesion": "sesión",
        "publica": "pública",
        "comunidad": "comunidad",
    }

    texto = texto.strip()

    for mal, bien in correcciones.items():
        texto = re.sub(rf"\b{mal}\b", bien, texto, flags=re.IGNORECASE)

    texto = re.sub(r"\s+", " ", texto)
    texto = texto.replace(" ,", ",").replace(" .", ".")
    texto = texto.replace(" :", ":").replace(" ;", ";")

    frases = re.split(r"(?<=[.!?])\s+", texto)
    frases_corregidas = []

    for frase in frases:
        frase = frase.strip()
        if frase:
            frase = frase[0].upper() + frase[1:]
            frases_corregidas.append(frase)

    texto = " ".join(frases_corregidas)

    reemplazos_logicos = {
        "dos mil veintiséis": "2026",
        "dos mil veintiseis": "2026",
        "tres de la tarde con diez minutos": "3:10 p. m.",
        "prestar toda la atención del caso": "brindar la atención correspondiente",
        "vamos a dar inicio con esta reunión": "se da inicio a la reunión",
    }

    for mal, bien in reemplazos_logicos.items():
        texto = re.sub(mal, bien, texto, flags=re.IGNORECASE)

    return texto


def dividir_en_parrafos(texto):
    oraciones = re.split(r"(?<=[.!?])\s+", texto.strip())
    parrafos = []
    bloque = []

    for oracion in oraciones:
        if oracion:
            bloque.append(oracion)
        if len(bloque) == 3:
            parrafos.append(" ".join(bloque))
            bloque = []

    if bloque:
        parrafos.append(" ".join(bloque))

    return parrafos


def buscar_tema(texto, tema):
    resultados = []
    if not texto or not tema:
        return resultados

    parrafos = dividir_en_parrafos(texto)

    for i, parrafo in enumerate(parrafos, start=1):
        if tema.lower() in parrafo.lower():
            resultados.append((i, parrafo))

    return resultados


def crear_word(nombre_asociacion, fecha_reunion, titulo_reunion, texto):
    documento = Document()

    section = documento.sections[0]
    section.top_margin = Pt(50)
    section.bottom_margin = Pt(50)
    section.left_margin = Pt(50)
    section.right_margin = Pt(50)

    titulo = documento.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run("ACTA / RESUMEN DE REUNIÓN")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(31, 78, 121)

    subtitulo = documento.add_paragraph()
    subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitulo.add_run(nombre_asociacion.upper())
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(84, 130, 53)

    documento.add_paragraph("")

    datos = documento.add_paragraph()
    datos.add_run("Título de la reunión: ").bold = True
    datos.add_run(titulo_reunion)

    datos = documento.add_paragraph()
    datos.add_run("Fecha de la reunión: ").bold = True
    datos.add_run(str(fecha_reunion))

    datos = documento.add_paragraph()
    datos.add_run("Fecha de generación del documento: ").bold = True
    datos.add_run(datetime.now().strftime("%d/%m/%Y"))

    documento.add_paragraph("")

    encabezado = documento.add_paragraph()
    run = encabezado.add_run("Desarrollo de la reunión")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(31, 78, 121)

    parrafos = dividir_en_parrafos(texto)

    for i, parrafo in enumerate(parrafos, start=1):
        p = documento.add_paragraph()
        p.style = "List Number"
        run = p.add_run(parrafo)
        run.font.size = Pt(11)

    documento.add_paragraph("")

    cierre = documento.add_paragraph()
    cierre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cierre.add_run("Documento generado automáticamente a partir de la transcripción de audio.")
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(100, 100, 100)

    archivo = io.BytesIO()
    documento.save(archivo)
    archivo.seek(0)
    return archivo


# =========================
# INTERFAZ
# =========================

st.title("🎙️ Transcriptor Inteligente de Reuniones")
st.write("Grabe audio de hasta **10 minutos**, transcriba, corrija el texto, busque temas y descargue en Word.")

with st.sidebar:
    st.header("📌 Datos del documento")

    nombre_asociacion = st.text_input(
        "Nombre de la asociación",
        value="Asociación de Desarrollo"
    )

    titulo_reunion = st.text_input(
        "Nombre o tema de la reunión",
        value="Reunión ordinaria"
    )

    fecha_reunion = st.date_input(
        "Fecha de la reunión",
        value=datetime.now()
    )

# =========================
# GRABADOR
# =========================

st.subheader("🎧 Grabación de audio")

audio = mic_recorder(
    start_prompt="🎙️ Iniciar grabación",
    stop_prompt="⏹️ Detener grabación",
    just_once=False,
    use_container_width=True,
    key="grabador_audio"
)

if audio:
    audio_bytes = audio["bytes"]

    st.audio(audio_bytes, format="audio/wav")

    duracion_estimada = len(audio_bytes) / 16000 / 2

    if duracion_estimada > LIMITE_SEGUNDOS:
        st.error(f"El audio supera el límite permitido de {LIMITE_MINUTOS} minutos.")
    else:
        if st.button("📝 Transcribir audio"):
            with st.spinner("Transcribiendo audio..."):

                with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as temp_audio:
                    temp_audio.write(audio_bytes)
                    temp_audio_path = temp_audio.name

                try:
                    resultado = modelo.transcribe(
                        temp_audio_path,
                        language="es",
                        fp16=False
                    )

                    texto = resultado["text"].strip()

                    if texto:
                        st.session_state.transcripcion_total += texto + "\n\n"
                        st.success("Audio transcrito correctamente.")
                    else:
                        st.warning("No se detectó texto.")

                except Exception as e:
                    st.error(f"Ocurrió un error al transcribir: {e}")

                finally:
                    if os.path.exists(temp_audio_path):
                        os.remove(temp_audio_path)

# =========================
# TEXTO ORIGINAL
# =========================

st.subheader("📄 Texto transcrito original")

texto_original = st.text_area(
    "Texto original:",
    value=st.session_state.transcripcion_total,
    height=250
)

st.session_state.transcripcion_total = texto_original

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("✨ Corregir texto"):
        st.session_state.texto_corregido = corregir_texto_basico(
            st.session_state.transcripcion_total
        )
        st.success("Texto corregido.")

with col2:
    if st.button("🧹 Limpiar todo"):
        st.session_state.transcripcion_total = ""
        st.session_state.texto_corregido = ""
        st.rerun()

with col3:
    st.download_button(
        label="⬇️ Descargar texto TXT",
        data=st.session_state.transcripcion_total,
        file_name="transcripcion_original.txt",
        mime="text/plain"
    )

# =========================
# TEXTO CORREGIDO
# =========================

st.subheader("✅ Texto corregido y editable")

texto_corregido = st.text_area(
    "Puede revisar, ampliar o corregir manualmente:",
    value=st.session_state.texto_corregido,
    height=300
)

st.session_state.texto_corregido = texto_corregido

# =========================
# BÚSQUEDA POR TEMA
# =========================

st.subheader("🔎 Búsqueda por tema")

tema_busqueda = st.text_input(
    "Escriba el tema a buscar. Ejemplo: contaminación del agua"
)

if st.button("Buscar tema"):
    resultados = buscar_tema(
        st.session_state.texto_corregido or st.session_state.transcripcion_total,
        tema_busqueda
    )

    if resultados:
        st.success(f"Se encontraron {len(resultados)} coincidencias.")
        for numero, parrafo in resultados:
            st.markdown(f"**Punto / párrafo {numero}:**")
            st.info(parrafo)
    else:
        st.warning("No se encontraron menciones sobre ese tema.")

# =========================
# DESCARGA WORD
# =========================

st.subheader("📘 Descargar documento Word")

texto_para_word = st.session_state.texto_corregido or st.session_state.transcripcion_total

if texto_para_word.strip():
    archivo_word = crear_word(
        nombre_asociacion=nombre_asociacion,
        fecha_reunion=fecha_reunion,
        titulo_reunion=titulo_reunion,
        texto=texto_para_word
    )

    st.download_button(
        label="⬇️ Descargar Word",
        data=archivo_word,
        file_name="reunion_transcrita.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.warning("Primero debe existir una transcripción o texto corregido para generar el Word.")

st.caption("Versión demo: transcripción de audio, corrección básica, búsqueda por tema y descarga en Word.")
