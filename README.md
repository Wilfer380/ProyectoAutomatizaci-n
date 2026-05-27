# Generador de Etiquetas SAP

Aplicación de escritorio en Python con PySide6 para generar, previsualizar e imprimir etiquetas de inventario desde un archivo Excel. La impresión objetivo es la **SATO WS408** con etiquetas físicas de **48 mm x 23 mm**.

## Quick path

1. Instalá dependencias con `pip install -r requirements.txt`.
2. Ejecutá `python main.py`.
3. Seleccioná el Excel de inventario.
4. Generá las etiquetas y revisalas en la vista previa nativa.
5. Confirmá para imprimir en **SATO WS408** o usá **Rehacer** para ajustar antes de imprimir.

## Requisitos de producción

| Requisito | Detalle |
|---|---|
| Sistema | Windows |
| Python | 3.11 o superior recomendado |
| Excel fuente | Archivo `.xlsx/.xlsm/.xls` con hoja `Hoja1` y columnas esperadas |
| Impresora | **SATO WS408** instalada en Windows con ese nombre exacto |
| Driver | Controlador/Printer Utility de **SATO WS4** instalado por TI/informática si el usuario no tiene permisos |

> Si el programa no detecta la SATO WS408, mostrará un mensaje claro pidiendo instalar el controlador o contactar a TI/informática.

## Arquitectura actual

```text
ProyectoAutomatizaci-n/
├─ main.py
├─ requirements.txt
├─ ui/
│  ├─ main_window.py
│  └─ preview_subwindow.py
├─ view_models/
│  ├─ main_view_model.py
│  ├─ label_item_view_model.py
│  └─ preview_view_model.py
├─ services/
│  ├─ excel_service.py
│  ├─ print_service.py
│  ├─ driver_check.py
│  └─ validation_service.py
├─ models/
│  ├─ asset_record.py
│  └─ app_settings.py
└─ deploy/
   ├─ installer_window_generadoretiquetassap.py
   ├─ launcher_generadoretiquetassap.py
   └─ printer_driver_preflight.py
```

## Decisiones técnicas

| Área | Decisión |
|---|---|
| UI | PySide6 con patrón **MVVM**. La vista no contiene lógica de negocio. |
| No bloqueo | La lectura/procesamiento corre con `QThread` para evitar congelar la pantalla. |
| Entrada | El **Excel es la única fuente**. Ya no se usa plantilla Word. |
| Extracción | `openpyxl` lee datos e imágenes ancladas del Excel. |
| Preview | Subventana PySide6 con escena lógica de etiqueta **48x23 mm**. |
| Impresión | `QPrinter` + `QPainter`, sin Word, sin COM y sin aplicaciones externas. |
| Driver | `QPrinterInfo` valida que exista **SATO WS408** antes de imprimir. |

## Instalación para desarrollo

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

## Validación

```powershell
python -m unittest discover -s tests
```

## Empaquetado

El empaquetado empresarial usa los `.spec` del repositorio:

```powershell
python deploy/build_release_generadoretiquetassap.py
```

El instalador valida el paquete, copia la aplicación, crea el launcher y advierte si no detecta el driver **SATO WS408**. Si el driver no está instalado, la instalación de la app puede terminar, pero la impresión real requerirá que TI/informática instale y configure el controlador.

## Estado del refactor SDD

- ✅ PR 1: extracción nativa desde Excel con `openpyxl` y eliminación de Word/COM.
- ✅ PR 2: ViewModels MVVM y procesamiento no bloqueante.
- ✅ PR 3: preview nativo en PySide6.
- ✅ PR 4: impresión directa con `QPrinter`.
- ✅ PR 5: detección de driver, mensajes amigables, instalador/launcher más robustos y limpieza de rutas legacy.
