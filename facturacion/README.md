# Importación de detalle de compras a EFFI

El conversor toma XML UBL/DIAN de facturas de proveedores, incluidos los XML que Effi entrega dentro de un `AttachedDocument`, y crea un Excel por factura usando la plantilla de importación de conceptos de compra.

## Estructura y configuración local

- `generar_importacion.py` convierte facturas XML en libros Excel para importación manual.
- `azure_codes.py` consulta códigos y productos en Azure SQL; ambas herramientas viven en esta integración y comparten las dependencias de `requirements.txt`, en la raíz del repositorio.
- `files/xml/` contiene entradas, `files/template/` la plantilla EFFI y `output/` los Excel generados localmente.
- `facturacion/.env.example` es la plantilla versionable de configuración. Cópiala como `facturacion/.env` y completa las credenciales; el archivo `.env` está excluido de Git.
- Los entornos virtuales `.venv*`, los CSV exportados y `output/` también se excluyen de Git. Crea el entorno virtual en cada máquina y reinstala las dependencias con `python -m pip install -r requirements.txt` desde la raíz del repositorio.

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

## Consulta de códigos en Azure SQL

El módulo `azure_codes.py` consulta `[stg].[ProductoVerificador]` para listar códigos de barras o buscar un producto. Requiere `unixODBC` y el Microsoft ODBC Driver for SQL Server indicado en `SQL_DRIVER`.

Instala las dependencias desde la raíz del repositorio:

```bash
python -m pip install -r requirements.txt
```

Completa `facturacion/.env` con `SQL_SERVER`, `SQL_DATABASE` y el método de autenticación. Para autenticación SQL usa `SQL_USERNAME` y `SQL_PASSWORD`; alternativamente configura `SQL_AUTHENTICATION`. Ajusta `SQL_DRIVER` al controlador instalado. `SQL_ENCRYPT`, `SQL_TRUST_CERT` y `SQL_CONNECTION_TIMEOUT` controlan el cifrado, la validación del certificado y el tiempo de espera.

Desde la raíz del repositorio, consulta todos los códigos o busca uno específico:

```bash
python -m facturacion.azure_codes
python -m facturacion.azure_codes --output codigos.csv
python -m facturacion.azure_codes --code 7701234567890
```

La conexión requiere que el firewall del servidor Azure SQL permita la IP de origen.

## Pruebas

Desde la raíz del repositorio:

```powershell
python -m unittest discover -s facturacion/tests -v
```