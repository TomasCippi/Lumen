import json
from pathlib import Path

def _obtener_carpeta_documentos() -> Path:
    """Devuelve la carpeta 'Documentos' real del usuario en Windows,
    incluso si está redirigida (ej. a OneDrive)."""
    try:
        import ctypes.wintypes
        CSIDL_PERSONAL = 5  # constante de Windows para "Mis Documentos"
        SHGFP_TYPE_CURRENT = 0
        buffer = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buffer)
        return Path(buffer.value)
    except Exception:
        # Fallback para no-Windows o si algo falla: usamos el estándar
        return Path.home() / "Documents"


CARPETA_CONFIG = _obtener_carpeta_documentos() / "Lumen"
ARCHIVO_CONFIG = CARPETA_CONFIG / "config.json"


def _leer_config() -> dict:
    """Lee el JSON completo. Si no existe todavía, devuelve un dict vacío."""
    if not ARCHIVO_CONFIG.exists():
        return {}

    try:
        with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        # Archivo corrupto o ilegible: arrancamos de cero en vez de romper la app
        return {}


def _escribir_config(data: dict) -> None:
    """Escribe el dict completo al JSON, creando la carpeta si hace falta."""
    CARPETA_CONFIG.mkdir(parents=True, exist_ok=True)
    with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def obtener_valor(clave: str, default=None):
    """Devuelve el valor guardado para esa clave, o `default` si no existe."""
    config = _leer_config()
    return config.get(clave, default)


def guardar_valor(clave: str, valor) -> None:
    """Guarda (o actualiza) un valor puntual, sin borrar el resto del config."""
    config = _leer_config()
    config[clave] = valor
    _escribir_config(config)


def eliminar_valor(clave: str) -> None:
    """Elimina una clave del config, si existe."""
    config = _leer_config()
    if clave in config:
        del config[clave]
        _escribir_config(config)