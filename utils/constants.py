import os
from pathlib import Path

from utils.runtime import get_user_home_dir


APP_NAME = "Generador de Etiquetas SAP"
APP_ORGANIZATION = "CO_MDE_DISENO_DI"
APP_SETTINGS_FILE = "app_settings.json"

DEFAULT_NETWORK_ROOT = Path(
    r"\\comde019\DFSMDE\PUBLIC\CO_MDE_DISENO_DI\RESPALDO DISEÑOS\SAP - Respaldo diseños\FORMATOS SAP\Automatización"
)
LOCAL_APPDATA_ROOT = Path(os.getenv("APPDATA", str(get_user_home_dir()))) / APP_ORGANIZATION / APP_NAME
LOGS_DIR_NAME = "logs"

EXCEL_SHEET_SOURCE = "Hoja1"
EXCEL_SHEET_LABEL = "Etiqueta provisional"

SOURCE_HEADERS = {
    "asset_id": "Activo fijo",
    "asset_name": "Denominación del activo fijo",
    "section": "Seccion",
}

LABEL_COLUMNS = {
    "asset": "J",
    "label": "K",
    "section": "L",
}
LABEL_OUTPUT_START_ROW = 2

BLOCK_SIZE = 27
TARGET_PRINTER_NAME = "SATO WS408"

# Nombres reales esperados de los grupos Excel que representan cada etiqueta visual.
EXCEL_LABEL_GROUP_NAMES = [
    "Group 18", "Group 19", "Group 26", "Group 90", "Group 97", "Group 104",
    "Group 111", "Group 118", "Group 125", "Group 190", "Group 197", "Group 204",
    "Group 211", "Group 218", "Group 225", "Group 232", "Group 239", "Group 246",
    "Group 253", "Group 260", "Group 267", "Group 274", "Group 281", "Group 288",
    "Group 295", "Group 302", "Group 309",
]

TEMP_IMAGES_DIR_NAME = "temp_images"

# Render de PNG suficientemente grande para que Word lo reduzca sin perder legibilidad.
LABEL_IMAGE_WIDTH_PX = 412
LABEL_IMAGE_HEIGHT_PX = 210

# Tamaño final DENTRO de la plantilla Word real.
WORD_LABEL_WIDTH_CM = 3.92
WORD_LABEL_HEIGHT_CM = 2.00

PROCESS_STATUS_IDLE = "idle"
PROCESS_STATUS_RUNNING = "running"
PROCESS_STATUS_ERROR = "error"
PROCESS_STATUS_COMPLETED = "completed"
