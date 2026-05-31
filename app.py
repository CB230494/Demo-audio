import streamlit as st
import tempfile
import os
import whisper
from streamlit_mic_recorder import mic_recorder

# =========================
# CONFIGURACIÓN GENERAL
# =========================

st.set_page_config(
    page_title="Transcriptor de Audio",
    page_icon="🎙️",
    layout="centered"
)

LIMITE_MINUTOS = 10
LIMITE_SEGUNDOS = LIMITE_MINUTOS * 60

# =========================
# CARGA DEL MODELO
# =========================

@st.cache_resource
def cargar_modelo():
    return whisper.load_model("base")

modelo = cargar_modelo()

# =========================
# ESTADO DE SESIÓN
# =========================

if "transcripcion_total" not in st.session_state:
    st.session_state.transcripcion_total = ""

# =========================
# INTERFAZ
# =========================

st.title("🎙️ Transcriptor de Audio a Texto")
st.write("Grabe un audio de hasta **10 minutos** y la app lo convertirá a texto.")

st.info(
    "Recomendación: hable claro, cerca del micrófono y evite mucho ruido de fondo."
)

audio = mic_recorder(
    start_prompt="🎙️ Iniciar grabación",
    stop_prompt="⏹️ Detener grabación",
    just_once=False,
    use_container_width=True,
    key="grabador_audio"
)

# =========================
# PROCESAR AUDIO GRABADO
# =========================

if audio:
    audio_bytes = audio["bytes"]

    st.audio(audio_bytes, format="audio/wav")

    duracion_estimada = len(audio_bytes) / 16000 / 2

    if duracion_estimada > LIMITE_SEGUNDOS:
        st.error(f"El audio supera el límite permitido de {LIMITE_MINUTOS} minutos.")
    else:
        if st.button("📝 Transcribir audio"):
            with st.spinner("Transcribiendo audio, espere un momento..."):

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
                        st.warning("No se detectó texto en el audio.")

                except Exception as e:
                    st.error(f"Ocurrió un error al transcribir: {e}")

                finally:
                    if os.path.exists(temp_audio_path):
                        os.remove(temp_audio_path)

# =========================
# RESULTADO
# =========================

st.subheader("📄 Texto transcrito")

texto_editado = st.text_area(
    "Puede revisar o corregir el texto:",
    value=st.session_state.transcripcion_total,
    height=300
)

st.session_state.transcripcion_total = texto_editado

col1, col2 = st.columns(2)

with col1:
    if st.button("🧹 Limpiar texto"):
        st.session_state.transcripcion_total = ""
        st.rerun()

with col2:
    st.download_button(
        label="⬇️ Descargar TXT",
        data=st.session_state.transcripcion_total,
        file_name="transcripcion_audio.txt",
        mime="text/plain"
    )

st.caption("Demo básica para audios de hasta 10 minutos.")



