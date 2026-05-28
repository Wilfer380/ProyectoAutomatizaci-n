# Guía de impresión y configuración de máquina — SATO WS408

Esta guía separa lo que controla la aplicación de lo que controla Windows/la impresora. Es clave para diagnosticar si un problema viene del software o de la configuración física de la máquina.

## Regla principal

Si la previsualización se ve correcta pero el papel sale corrido, cortado, muy claro o no sale, el problema normalmente está en **driver, Windows, spooler, sensor, calibración o medio físico**, no en el layout de la app.

## Qué controla la aplicación

| Control | Valor aplicado |
|---|---|
| Impresora destino | `SATO WS408` |
| Tamaño lógico | `48 mm x 23 mm` |
| Resolución | `203 DPI` |
| Márgenes enviados | `0 mm` |
| Render | `LabelRenderer` compartido entre preview e impresión |
| Lote | Un trabajo de impresión independiente por etiqueta (`separate_jobs=True`) |
| Seguridad | Preview obligatorio antes de imprimir |

## Qué controla la máquina / Windows

| Área | Debe validarse en la máquina |
|---|---|
| Driver | Driver/Printer Utility compatible con SATO WS4 instalado. |
| Nombre | La cola debe llamarse exactamente `SATO WS408`. |
| Stock/tamaño | Papel personalizado de `48 mm x 23 mm`. |
| Sensor | Gap/label sensor calibrado según el rollo instalado. |
| Calibración | La impresora debe reconocer el inicio y final de cada etiqueta. |
| Oscuridad/velocidad | Ajuste físico/driver para que el texto salga legible sin quemarse. |
| Spooler | Cola de Windows sin trabajos atascados. |
| Escalado | No usar “fit to page”, “ajustar a página” ni escalado automático. |

## Checklist para TI / informática

1. Instalar el driver oficial o compatible **SATO WS4**.
2. Crear/verificar la impresora en Windows con nombre exacto: `SATO WS408`.
3. En preferencias de impresión, crear o seleccionar stock personalizado:
   - ancho: `48 mm`;
   - alto: `23 mm`;
   - márgenes: `0 mm` si el driver lo permite;
   - resolución: `203 DPI`.
4. Verificar que el tipo de medio sea etiqueta con separación/gap, no papel continuo, si aplica al rollo.
5. Calibrar sensor/medio desde SATO Printer Utility o desde el panel/botones de la impresora.
6. Limpiar cola de impresión de Windows antes de pruebas grandes.
7. Imprimir una prueba de hardware/driver si la utilidad SATO lo permite.
8. Abrir la app e imprimir primero **1 etiqueta**.
9. Si una etiqueta sale bien, probar un lote pequeño de **3 a 5 etiquetas**.
10. Si el lote pequeño sale estable, recién ahí imprimir el lote completo.

## Diagnóstico rápido

| Síntoma | Probable origen | Qué hacer |
|---|---|---|
| La app dice que no detecta impresora | Nombre o driver no instalado. | Renombrar cola a `SATO WS408` o reinstalar driver. |
| En preview se ve mal | Datos/layout/app. | Revisar Excel, imagen anclada o `LabelRenderer`. |
| Preview bien, papel corrido horizontal/vertical | Stock, origen, margen o calibración del driver. | Revisar tamaño 48x23, márgenes y calibración. |
| Primera etiqueta bien, siguientes se corren | Driver/spooler/origen de página. | Mantener `separate_jobs=True`; limpiar cola; revisar driver. |
| Sale muy claro u oscuro | Oscuridad, velocidad, ribbon/medio. | Ajustar darkness/speed en driver/utilidad SATO. |
| No sale nada pero no hay error en app | Spooler, impresora pausada/offline o driver. | Revisar cola de Windows y estado físico. |
| Corta en medio de etiqueta | Sensor/gap no calibrado o stock incorrecto. | Recalibrar medio y confirmar 48x23 mm. |

## Criterio de aceptación físico

Antes de liberar una máquina para producción, guardar esta evidencia:

- [ ] Screenshot de Windows mostrando la cola `SATO WS408`.
- [ ] Screenshot o nota de preferencias con stock `48 mm x 23 mm`.
- [ ] Prueba física de 1 etiqueta correcta.
- [ ] Prueba física de 3–5 etiquetas sin corrimiento acumulado.
- [ ] Confirmación de que el usuario sabe usar preview y rechazar si algo no se ve bien.

## Cuándo tocar código y cuándo no

| Caso | Acción correcta |
|---|---|
| Preview y papel están mal igual | Revisar app/layout/datos. |
| Preview está bien pero papel sale mal | Revisar máquina/driver/calibración. |
| Solo falla una máquina | No cambiar layout global; corregir esa máquina. |
| Falla en todas las máquinas con el mismo patrón | Revisar `LabelRenderer`, DPI o tamaño lógico. |

## Recomendación operativa

Mantener una etiqueta física aprobada como muestra patrón. Cuando se instala una nueva máquina, comparar contra esa muestra antes de imprimir lotes grandes.
