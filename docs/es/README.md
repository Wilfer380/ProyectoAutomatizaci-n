# Generador de etiquetas SATO WS408 — Documentación técnica y operativa

Esta aplicación de escritorio permite cargar un Excel de inventario, seleccionar filtros y etiquetas, previsualizar cada etiqueta como se imprimirá y enviar el lote confirmado directamente a una impresora **SATO WS408** con etiquetas físicas de **48 mm x 23 mm**.

## Camino rápido

1. Instalá dependencias: `pip install -r requirements.txt`.
2. Ejecutá: `python main.py`.
3. Seleccioná el archivo Excel de inventario.
4. Abrí **Seleccionar filtros**.
5. Marcá filtros o hacé doble click en un filtro para elegir etiquetas específicas.
6. Presioná **Generar etiquetas**.
7. Revisá la previsualización completa con scroll.
8. Confirmá con **Confirmar e imprimir** o cancelá con **Rechazar**.

## Qué hace la aplicación

| Etapa | Comportamiento |
|---|---|
| Entrada | Lee un Excel desde `Hoja1`; la columna `Seccion` alimenta los filtros. |
| Selección | Permite seleccionar filtros completos o etiquetas individuales dentro de un filtro. |
| Previsualización | Renderiza todas las etiquetas seleccionadas en una ventana blanca con scroll. |
| Aprobación | El usuario decide si imprime todo el lote o rechaza/cancela. |
| Impresión | Envía cada etiqueta a `SATO WS408` usando `QPrinter` y `QPainter`. |
| Errores | Muestra mensajes amigables cuando falta el driver o el Excel no se puede leer. |

## Qué no hace

- No usa Microsoft Word.
- No usa automatización COM para generar, previsualizar ni imprimir etiquetas.
- No imprime sin aprobación visual previa.
- No instala automáticamente el driver de la SATO; guía al usuario a contactar TI/informática.

## Arquitectura

La app sigue un patrón **MVVM** simple:

| Capa | Responsabilidad | Archivos principales |
|---|---|---|
| UI / Views | Widgets, diálogos y eventos visuales. | `ui/main_window.py`, `ui/filter_selection_dialog.py`, `ui/label_selection_dialog.py`, `ui/preview_subwindow.py` |
| ViewModels | Estado de pantalla, señales y conversión a modelos de etiqueta. | `view_models/main_view_model.py`, `view_models/preview_view_model.py`, `view_models/label_item_view_model.py` |
| Services | Lectura de Excel, render, impresión y validación del driver. | `services/excel_service.py`, `services/print_service.py`, `services/driver_check.py` |
| Models | Datos puros de negocio/configuración. | `models/asset_record.py`, `models/app_settings.py` |
| Deploy | Instalador, launcher y preflight de impresora. | `deploy/` |

Diagrama: [`../shared/diagrams/c4-container.mmd`](../shared/diagrams/c4-container.mmd)

## Flujo MVVM

1. La vista emite eventos: seleccionar Excel, seleccionar filtros, generar etiquetas.
2. `MainViewModel` carga registros, agrupa filtros y mantiene selección.
3. La vista recibe señales (`filtersLoaded`, `recordCountChanged`, `processingFinished`).
4. `PreviewViewModel` recibe las etiquetas listas para revisar.
5. La subventana de preview llama callbacks inyectados para renderizar e imprimir.

Diagrama: [`../shared/diagrams/mvvm-signal-flow.mmd`](../shared/diagrams/mvvm-signal-flow.mmd)

## Extracción desde Excel

La fuente oficial es el Excel. El contrato esperado es:

| Dato | Fuente |
|---|---|
| ID del activo | Columna `Activo fijo` |
| Descripción | Columna `Denominación del activo fijo` |
| Filtro/sección | Columna `Seccion` de `Hoja1` |
| Imagen/logo | Imagen anclada en la misma fila del registro |

`ExcelService` usa `openpyxl` para abrir el archivo sin depender de Excel instalado. Las imágenes se mapean por la fila de anclaje y se convierten a `QImage` para que el mismo objeto pueda usarse en preview e impresión.

Diagrama: [`../shared/diagrams/excel-extraction-flow.mmd`](../shared/diagrams/excel-extraction-flow.mmd)

## Selección de filtros y etiquetas

La selección está diseñada para evitar imprimir etiquetas incorrectas:

1. El usuario abre **Seleccionar filtros**.
2. La app muestra cada valor de `Seccion` con la cantidad de etiquetas asociadas.
3. Si el usuario marca un filtro sin abrir detalle, se imprimen todas sus etiquetas.
4. Si hace doble click en el filtro, abre una subventana con checklist por etiqueta.
5. La app imprime solo las etiquetas seleccionadas explícitamente.

Diagrama: [`../shared/diagrams/filter-label-checklist-flow.mmd`](../shared/diagrams/filter-label-checklist-flow.mmd)

## Previsualización y aprobación

La ventana de previsualización es una barrera de seguridad antes de imprimir:

- Fondo blanco para simular la etiqueta física.
- Scroll vertical para revisar lotes grandes, por ejemplo 179 etiquetas.
- Muestra varias etiquetas visibles a la vez.
- Usa el mismo `LabelRenderer` que la impresión real.
- **Confirmar e imprimir** envía todo el lote a la impresora.
- **Rechazar** cancela el proceso sin imprimir.

Diagrama: [`../shared/diagrams/preview-approval-state.mmd`](../shared/diagrams/preview-approval-state.mmd)

## Impresión SATO WS408

Configuración aprobada:

| Parámetro | Valor |
|---|---|
| Impresora | `SATO WS408` |
| Tamaño físico | `48 mm x 23 mm` |
| Resolución | `203 DPI` |
| Tamaño renderizado | aprox. `384 x 184 px` |
| Márgenes | `0 mm` |
| Estrategia de lote | un trabajo independiente por etiqueta (`separate_jobs=True`) |

La impresión se hace con `QPrinter` y `QPainter`, sin Word ni aplicaciones externas. Cada etiqueta se manda como trabajo independiente porque en pruebas reales esa estrategia estabilizó la posición y evitó corrimientos acumulados del driver.

Diagrama: [`../shared/diagrams/sato-print-pipeline.mmd`](../shared/diagrams/sato-print-pipeline.mmd)

## Decisiones y por qué se tomaron

| Decisión | Por qué |
|---|---|
| Eliminar Word/COM | Word fallaba entre máquinas, agregaba dependencia externa y hacía difícil previsualizar/imprimir de forma consistente. |
| Usar `openpyxl` | Permite leer Excel e imágenes sin abrir Excel; mejora testabilidad y portabilidad dentro de Windows. |
| Usar PySide6 nativo | Unifica UI, preview e impresión en el mismo runtime. |
| Adoptar MVVM | Reduce lógica en widgets y separa estado, eventos y servicios. |
| Compartir renderer entre preview e impresión | Evita que el usuario apruebe una imagen distinta a la que se imprime. |
| Validar driver antes de imprimir | Falla temprano con mensaje accionable para usuario/TI. |
| Imprimir en trabajos separados | Estabiliza la SATO WS408 y evita desplazamientos en lotes. |
| Checklist por filtro y etiqueta | Minimiza errores humanos cuando un filtro contiene muchas etiquetas. |

## Buenas prácticas aplicadas

- Separación de responsabilidades entre UI, ViewModels, servicios y modelos.
- Inyección de callbacks para preview/impresión, facilitando pruebas.
- Señales/slots de PySide6 para evitar acoplar directamente widgets con servicios.
- Mensajes de error orientados a acción, no trazas técnicas para usuarios finales.
- Pruebas unitarias para extracción, ViewModels, preview, driver y servicio de impresión.
- Previsualización obligatoria antes de imprimir.
- Configuración explícita de impresora, tamaño físico, DPI y márgenes.

## Manejo de errores

| Error | Respuesta esperada |
|---|---|
| Excel inexistente o inválido | Mostrar error amigable y no continuar. |
| No hay registros | No imprimir; informar que no hay etiquetas. |
| Driver SATO ausente | Pedir instalar/configurar `SATO WS408` o contactar TI. |
| Fallo al iniciar impresión | Mostrar error y permitir reintento. |
| Excepción inesperada | Guardar detalle técnico en logs y mostrar mensaje seguro. |

Diagrama: [`../shared/diagrams/error-handling-boundaries.mmd`](../shared/diagrams/error-handling-boundaries.mmd)

## Instalación y despliegue

El flujo de despliegue incluye un preflight del driver:

1. Se instala o copia la app.
2. El launcher/preflight verifica si Windows conoce la impresora `SATO WS408`.
3. Si falta el driver, la app puede quedar instalada, pero se informa que la impresión real requiere soporte de TI.
4. El usuario final no debe resolver manualmente drivers si no tiene permisos.

Diagrama: [`../shared/diagrams/deployment-driver-preflight.mmd`](../shared/diagrams/deployment-driver-preflight.mmd)

## Validación para mantenedores

Antes de abrir o mergear cambios:

```powershell
python -m unittest discover -s tests
```

Checklist de revisión:

- [ ] No se reintrodujo Word/COM.
- [ ] Preview e impresión siguen usando el mismo renderer.
- [ ] El tamaño sigue siendo 48 mm x 23 mm.
- [ ] La impresora objetivo sigue siendo `SATO WS408`.
- [ ] La selección por filtro/etiqueta no imprime registros no seleccionados.
- [ ] Los mensajes de error siguen siendo claros para usuarios no técnicos.

## Solución de problemas

| Síntoma | Causa probable | Acción |
|---|---|---|
| No aparece la impresora | Driver no instalado o nombre distinto. | Instalar/configurar `SATO WS408` desde TI. |
| El Excel no carga filtros | La columna `Seccion` no existe o está vacía. | Revisar `Hoja1` y encabezados. |
| La etiqueta se ve vacía | Falta información del activo o imagen mal anclada. | Revisar fila del Excel. |
| El lote imprime corrido | Cambios en driver/configuración. | Mantener `separate_jobs=True` y validar en SATO real. |
| Preview no coincide con impresión | Se modificó renderer o DPI. | Verificar `LabelRenderer` y `LabelPrintConfig`. |
