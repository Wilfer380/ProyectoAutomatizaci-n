import sys
from argparse import ArgumentParser
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMessageBox

from controllers.main_controller import MainController
from ui.main_window import MainWindow
from utils.config import ConfigManager
from utils.constants import APP_NAME, APP_ORGANIZATION
from utils.logger import install_global_exception_handlers, log_runtime_context, setup_logger
from utils.resources import resource_path


def _build_parser() -> ArgumentParser:
    parser = ArgumentParser(add_help=True)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--worker-action", choices=("load-filters", "run-process"), default="")
    parser.add_argument("--excel", default="")
    parser.add_argument("--word", default="")
    parser.add_argument("--filter", dest="selected_filter", default="")
    parser.add_argument("--printer", default="")
    parser.add_argument("--simulate", action="store_true")
    parser.add_argument("--working-directory", default="")
    parser.add_argument("--output-directory", default="")
    return parser


def _set_windows_app_id() -> None:
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            f"{APP_ORGANIZATION}.{APP_NAME}".replace(" ", ".")
        )
    except Exception:
        pass


def _build_app_icon() -> QIcon:
    candidates = (
        resource_path("assets", "app_icon.ico"),
        resource_path("assets", "app_icon.png"),
        resource_path("assets", "app_icon.svg"),
    )
    for candidate in candidates:
        if Path(candidate).exists():
            icon = QIcon(str(candidate))
            if not icon.isNull():
                return icon
    return QIcon()


def main() -> int:
    args, _ = _build_parser().parse_known_args()
    if args.worker:
        from services.worker_process import worker_main

        return worker_main(sys.argv[1:])

    logger = setup_logger()
    _set_windows_app_id()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(APP_ORGANIZATION)

    def show_unhandled_error(message: str) -> None:
        QMessageBox.critical(None, "Error en la aplicación", message)

    install_global_exception_handlers(logger, show_unhandled_error)
    log_runtime_context(logger)

    app_icon = _build_app_icon()
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    try:
        logger.info("Cargando configuración de usuario.")
        config_manager = ConfigManager()
        settings = config_manager.load()

        logger.info("Construyendo interfaz principal.")
        window = MainWindow(settings)
        if not app_icon.isNull():
            window.setWindowIcon(app_icon)
        window.controller = MainController(window, config_manager, logger)
        window.show()
        logger.info("Interfaz cargada correctamente.")

        exit_code = app.exec()
        logger.info("Cierre normal de aplicación. Código de salida: %s", exit_code)
        return exit_code
    except Exception:
        logger.exception("Error de arranque o ejecución principal.")
        QMessageBox.critical(
            None,
            "Error en la aplicación",
            "No se pudo iniciar o continuar la aplicación. El detalle técnico fue guardado en la carpeta logs.",
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
