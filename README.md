# Generador de Etiquetas SAP

Aplicación de escritorio en Python con PySide6 para automatizar la generación e impresión de etiquetas de inventario usando un archivo Excel, una plantilla Word y la impresora **SATO WS408**.

## Objetivo

La aplicación permite:

- seleccionar el archivo Excel origen,
- seleccionar la plantilla Word,
- cargar filtros únicos desde la columna **C** de **Hoja1**,
- procesar los registros filtrados en bloques de **27**,
- escribir los activos del bloque en **Etiqueta provisional!J2:J28**,
- validar los activos generados,
- copiar cada etiqueta individual desde los grupos/shapes de Excel,
- generar automáticamente un documento Word por bloque con **una etiqueta por página**,
- pegar cada etiqueta inline, centrada, con tamaño visual **5.03 cm x 2.54 cm**,
- validar y autoajustar posición/tamaño contra la celda Word,
- generar un documento Word nuevo por bloque copiando la plantilla original,
- imprimir automáticamente cada bloque en **SATO WS408**.

## Arquitectura

```text
Automatización/
├─ main.py
├─ requirements.txt
├─ README.md
├─ ui/
│  └─ main_window.py
├─ controllers/
│  └─ main_controller.py
├─ services/
│  ├─ excel_service.py
│  ├─ word_service.py
│  ├─ print_service.py
│  ├─ validation_service.py
│  └─ process_service.py
├─ models/
│  ├─ asset_record.py
│  └─ app_settings.py
├─ utils/
│  ├─ logger.py
│  ├─ config.py
│  └─ constants.py
└─ logs/
```

## Requisitos del entorno

- Windows
- Microsoft Excel instalado
- Microsoft Word instalado
- Impresora **SATO WS408** instalada con ese nombre exacto
- Python 3.11 o superior recomendado

## Instalación

### 1. Crear entorno virtual

```powershell
python -m venv .venv
```

### 2. Activar entorno virtual

```powershell
.venv\Scripts\Activate.ps1
```

### 3. Instalar dependencias

```powershell
pip install -r requirements.txt
```

## Ejecución

```powershell
python main.py
```

## Flujo funcional

1. Abrir la aplicación.
2. Seleccionar el archivo Excel.
3. Seleccionar la plantilla Word.
4. Pulsar **Actualizar filtros**.
5. Elegir un filtro del ComboBox.
6. Pulsar **Prueba visual sin imprimir** o **Generar e imprimir**.
7. La aplicación:
   - valida archivos, hojas, columnas e impresora,
   - filtra registros de **Hoja1** por la columna **C**,
   - procesa en bloques de **27**,
   - escribe activos filtrados en **Etiqueta provisional!J2:J28**,
   - valida activos,
   - copia las etiquetas desde shapes/grupos individuales de Excel,
   - copia la plantilla Word original para cada bloque,
   - crea un layout Word de una etiqueta por página sobre la copia del bloque,
   - inserta salto de página antes de cada etiqueta excepto la primera,
   - valida tamaño y contenedor de cada etiqueta, corrigiendo si excede,
   - en prueba visual genera documentos separados `{filtro}_bloque_001.docx`, `{filtro}_bloque_002.docx`, etc.; cada documento contiene hasta 27 páginas/etiquetas, abre el primero sin imprimir y deja el resto en la carpeta de simulación,
   - en impresión real imprime automáticamente cada copia de bloque de 27 páginas/etiquetas.

## Validaciones implementadas

- existencia del archivo Excel,
- existencia de la plantilla Word,
- presencia de hojas `Hoja1` y `Etiqueta provisional`,
- existencia de columnas:
  - `Activo fijo`
  - `Denominación del activo fijo`
  - `Seccion`
- filtro obligatorio,
- filtro con registros,
- impresora `SATO WS408` instalada,
- coincidencia de activos entre origen y salida,
- tamaño máximo de bloque de 27 activos, manteniendo una página Word por etiqueta,
- tamaño visual esperado de etiqueta **5.03 cm x 2.54 cm**,
- límites del contenedor Word con imagen inline centrada y una única etiqueta por página para evitar agrupamientos en una misma hoja.

## Decisiones técnicas clave

### Excel

Se usa **COM con pywin32** porque la hoja **Etiqueta provisional** tiene fórmulas y formato.  
Por eso la aplicación:

- escribe los activos del bloque únicamente en `Etiqueta provisional!J2:J28`,
- recalcula Excel,
- copia el resultado visual real desde shapes/grupos individuales.

### Word

Se usa **COM con pywin32** para generar el documento automático en Word.  
La plantilla original se mantiene intacta: cada bloque trabaja sobre una copia `.docx` propia. Sobre esa copia se limpia el contenido y se crea una página por etiqueta: antes de cada etiqueta posterior a la primera se inserta un salto de página, se agrega un contenedor 1x1 centrado y se pega la imagen inline con validación visual programática. Las constantes históricas de grilla 3x9 quedan solo como referencia legacy del bloque Excel; no gobiernan el flujo automático de Word.

### Impresión

Cada bloque se imprime automáticamente al terminar, sin confirmación intermedia.

## Advertencias técnicas

### 1. Word y Excel deben estar instalados

La solución depende del motor COM de Microsoft Office.  
No sirve en equipos sin Office.

### 2. La plantilla Word puede requerir ajuste fino

La implementación busca shapes por ID.  
Si la plantilla real fue modificada manualmente, los IDs o el comportamiento interno de los grupos pueden cambiar.

### 3. Archivos fuente

La aplicación trabaja sobre copias temporales o de simulación de Excel y Word para no sobrescribir los archivos fuente.  
Esto exige:

- permisos de escritura,
- archivo no bloqueado,
- espacio temporal disponible durante el proceso.

### 4. Red compartida

El proyecto trabaja sobre una ruta UNC compartida.  
Si la red está lenta o el archivo está bloqueado por otro usuario, el proceso puede fallar.

### 5. Impresora

La impresora debe existir con el nombre exacto:

```text
SATO WS408
```

## Empaquetado a EXE

Recomendación:

```powershell
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed main.py
```

Para una entrega empresarial más estable, probablemente convenga usar:

```powershell
pyinstaller --noconfirm --windowed --name GeneradorEtiquetasSAP main.py
```

## Próxima fase

Antes de usar en producción, hay que hacer pruebas reales con:

- el archivo Excel real,
- la plantilla Word real,
- la impresora real,
- los grupos reales de Excel usados para las etiquetas,
- la zona exacta `Etiqueta provisional!J2:J28`.

Ahí se ajustan los detalles finos que ninguna teoría reemplaza.
