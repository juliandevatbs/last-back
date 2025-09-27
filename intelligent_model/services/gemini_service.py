import os
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

# Get the protected api key
GEMINI_API_KEY = os.getenv("API_KEY")

# Verificar que la API key existe
if not GEMINI_API_KEY:
    raise ValueError("API_KEY no encontrada en el archivo .env")

# Configurar la API key
genai.configure(api_key="AIzaSyA4dk_DW2pGPZ89_GoE5tiKjSbu7DPBrpk")


def ask_gemini(prompt, model="gemini-1.5-flash"):
    try:
        # Crear el modelo
        print(f"USANDO MODELO {model}")

        model = genai.GenerativeModel(model)

        # Generar respuesta
        response = model.generate_content(prompt)

        return response.text

    except Exception as e:
        print(f"Error procesando documento: {e}")
        return None