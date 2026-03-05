# Cotizaciones JM-VISION (v2)

Archivo final: **`Cotizaciones_JMVISION_LIMPIO.xlsm`**

## Uso en 1 minuto
1. GitHub → **Code** → **Download ZIP**.
2. Abrir `Cotizaciones_JMVISION_LIMPIO.xlsm` en Excel.
3. Habilitar edición / contenido.
4. En hoja `COTIZACION`:
   - Selecciona `Documento` (COTIZACIÓN / NOTA DE VENTA).
   - Escribe códigos en columna `CODIGO` y cantidades.
   - Cambia `Imprimir con fotos` a `Sí/No` según necesidad.
5. Exporta con macro `ExportarPDFCarta`.

## Estructura
- `PRODUCTOS` (incluye columna `FOTO`, precios Final/Técnico y `MO-PUNTO`)
- `CLIENTES`
- `COTIZACION` (formato carta multi-página)
- `HISTORICO_COTIZACIONES`
- `VENTAS`
- `DASHBOARD`

## Macros (importables)
Se incluye `vba/JMVisionMacros.bas` con macros funcionales:
- `NuevaCotizacion`
- `GuardarEnHistorico`
- `ConvertirAVenta`
- `ExportarPDFCarta`
- `ActualizarCodigosYFotos`

Importar: `Alt+F11` → **File > Import File** → `vba/JMVisionMacros.bas`.

## Fotos de productos
- Columna `FOTO` se toma de `PRODUCTOS!D`.
- Ruta base configurable en `COTIZACION!N2` (por defecto `C:\JMVISION_FOTOS\`).
- El logo en encabezado usa `assets/JM-LOGO LARGO OFICIAL.png`.

## Datos de prueba
- Ejemplo corto (12 ítems) y ejemplo largo (24 ítems) registrados en `HISTORICO_COTIZACIONES`.
