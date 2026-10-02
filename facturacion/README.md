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
- Si una factura repite el mismo producto, consolida las unidades y calcula el costo efectivo como `neto total del grupo / unidades totales`. Esto distribuye el costo entre unidades pagadas y bonificadas sin cambiar el neto de factura. Requiere el mismo GTIN, `IdProducto`, referencia interna o descripción exacta (cuando no hay código), y la misma unidad; solo consolida si las líneas con importe comparten el perfil de impuesto. Perfiles fiscales incompatibles se mantienen separados.
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

Para consultar el catálogo EFFI y generar las excepciones de revisión:

```powershell
python -m facturacion.generar_importacion --azure-lookup
```

Si la factura incluye código de barras, el artículo se busca exclusivamente por ese código: no se asigna un producto por descripción cuando el GTIN no aparece en el catálogo. En ese caso se genera una alerta para vincular el código al producto desde EFFI. La búsqueda por descripción, marca, presentación y precios solo se usa cuando la línea no trae código de barras; las coincidencias dudosas quedan en `Excepciones` para revisión. Para los grupos con bonificación, la hoja incluye líneas XML de origen, unidades totales, neto agrupado y costo unitario efectivo. También registra diferencias entre el descuento indicado en el XML y el neto de línea; el Excel ajusta el descuento para que cantidad × precio - descuento reproduzca el neto declarado por la factura. La conexión requiere que el firewall de Azure SQL permita la IP de origen.

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

## Revisión humana de excepciones

Inicia el visor Streamlit desde la raíz del repositorio:

```powershell
streamlit run facturacion/revisar_excepciones.py
```

El visor muestra todas las facturas y productos que tienen XML en `facturacion/files/xml`. La descarga directa desde EFFI todavía no está integrada: para incorporar una factura se debe obtener su XML y colocarlo allí. Al abrir el visor, los XML nuevos o modificados se convierten automáticamente en Excel; **Sincronizar** vuelve a consultar Azure, asigna los GTIN únicos y regenera las facturas. Los códigos inexistentes o asociados a más de un artículo quedan como excepciones para decisión humana.

El menú lateral organiza **Facturas**, **Revisión de productos**, **Excepciones de GTIN** e **Historial**. La bandeja permite buscar y filtrar facturas; la cola de productos abre directamente la línea que requiere decisión, y el historial conserva las decisiones y aprobaciones. La sincronización trabaja con los XML locales y el catálogo Azure; no descarga facturas desde EFFI.

1. Opcionalmente, para generar los libros desde la terminal:

	```powershell
	python -m facturacion.generar_importacion --azure-lookup
	```

2. Inicia el visor:

	```powershell
	streamlit run facturacion/revisar_excepciones.py
	```

3. Abre la URL local que muestra Streamlit y corrige únicamente las líneas que aparecen como pendientes: GTIN sin match único, candidatos de baja confianza, cantidades inválidas o diferencias de neto. Los GTIN únicos y grupos compatibles del mismo producto se asignan automáticamente; las bonificaciones al 100% se detectan y su costo se prorratea como `neto pagado / unidades pagadas y bonificadas`. Si un vínculo o impuesto no es seguro, queda para decisión humana.

4. Cuando no queden excepciones pendientes, escribe el nombre de quien aprueba y usa **Generar Excel revisado**. Solo entonces se habilita la descarga del archivo `*_revisado.xlsx`. La aprobación registra responsable, fecha y una huella del XML y de las decisiones; si cambias una decisión o el XML, tendrás que aprobar de nuevo. Los costos, enlaces y aprobaciones quedan en `facturacion/output/revisiones_manuales.json`. Los archivos `*_effi.xlsx` generados antes de la aprobación son borradores: no los importes en EFFI.

En **Excepciones GTIN** aparecen los códigos no registrados junto con la factura y la descripción del XML. No se resuelven automáticamente por descripción: se deben vincular al producto correcto en EFFI. Para habilitar descarga automática de facturas se necesita integrar la API oficial de EFFI y configurar su autenticación; este proyecto no contiene actualmente ese conector.

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