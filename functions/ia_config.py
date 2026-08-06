from functions.config_manager import obtener_valor

CLAVE_API_KEY_GEMINI = "api_key_gemini"


def obtener_api_key_gemini():
    """Devuelve la API Key de Gemini guardada, o None si todavía no se configuró."""
    return obtener_valor(CLAVE_API_KEY_GEMINI)