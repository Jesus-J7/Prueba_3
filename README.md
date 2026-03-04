# Cotizaciones JM-VISION

Archivo principal: **`Cotizaciones_JMVISION.xlsm`**

## Uso en 1 minuto
1. GitHub → **Code** → **Download ZIP**.
2. Descomprime y abre `Cotizaciones_JMVISION.xlsm` en Microsoft Excel.
3. Habilita edición / contenido al abrir.
4. En `COTIZACION`:
   - Elige `TipoCliente` (Final o Técnico).
   - Escribe `CODIGO` y `CANTIDAD` en las líneas.
   - Verifica que se autocompletan descripción, precio y total.
5. Usa los botones visuales de la fila inferior como guía de flujo.

## Estructura incluida
- `PRODUCTOS` (base con productos y `MO-PUNTO`)
- `CLIENTES`
- `COTIZACION` (plantilla imprimible carta)
- `HISTORICO_COTIZACIONES`
- `VENTAS`
- `DASHBOARD`

## Nota técnica
- Se incluye `vba/JMVisionMacros.bas` con la base de macros para importar en el editor VBA de Excel (`Alt+F11` → Import File), en caso de querer vincular botones ActiveX/Form Controls a estas rutinas.
- El libro está en formato `.xlsm` para compatibilidad de macros y distribución por ZIP.
