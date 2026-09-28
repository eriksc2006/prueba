# Importación de detalle de compras a EFFI

El conversor toma XML UBL/DIAN de facturas de proveedores, incluidos los XML que Effi entrega dentro de un `AttachedDocument`, y crea un Excel por factura usando la plantilla de importación de conceptos de compra.

## Alcance de esta versión

- Solo diligencia el detalle de productos; no automatiza ni completa el encabezado de la factura.
- En **Artículo** usa, en orden, el identificador asignado por el comprador, el identificador estándar (normalmente GTIN) y el código del proveedor.
- Lleva cantidad, precio unitario, descuento total de línea y código de impuesto configurado.
- Para líneas promocionales con precio XML cero y descuento del 100%, reconstruye el precio bruto desde el descuento; corrige diferencias de redondeo de hasta $0,10 por línea para que el detalle importado reproduzca el neto XML.
- Deja vacíos tipo de egreso, lote, serie y observación porque no se pueden inferir de forma fiable del detalle UBL de los ejemplos.
- La tasa del 19% se asocia inicialmente al código EFFI `1`, mostrado en una fila de ejemplo de la plantilla. Verifica ese código en la empresa de EFFI antes de importar. Las demás tasas quedan sin código hasta configurarlas.
- No inicia sesión ni crea facturas en Effi; genera archivos para revisión e importación manual.

## Uso

Instala las dependencias desde la raíz del repositorio:

```powershell
python -m pip install -r requirements.txt
```

Procesa todos los XML de ejemplo (o los archivos que se coloquen en esa carpeta):

```powershell
python -m facturacion.generar_importacion
```

Indica otra carpeta/archivo y otro destino:

```powershell
python -m facturacion.generar_importacion --input "D:\\facturas\\xml" --output-dir "D:\\facturas\\effi"
```

Configura códigos tributarios adicionales con el código vigente en EFFI:

```powershell
python -m facturacion.generar_importacion --tax-code 5=3 --tax-code 0=4
```

Para guardar junto al Excel una representación JSON de los datos extraídos:

```powershell
python -m facturacion.generar_importacion --json
```

El programa muestra avisos si faltan identificadores de artículo, hay tasas tributarias sin mapear o los valores de precio, cantidad, descuento y neto no concuerdan. Revisa esos avisos antes de importar en EFFI.

## Pruebas

Desde la raíz del repositorio:

```powershell
python -m unittest discover -s facturacion/tests -v
```