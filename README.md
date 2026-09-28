# Integraciones con EFFI

Este repositorio será el espacio central para los desarrollos que impliquen integraciones con **EFFI**. Su propósito es organizar, documentar y mantener en un solo lugar las distintas automatizaciones, procesos y herramientas que intercambien información con ese sistema.

## Organización del proyecto

Cada integración debe desarrollarse dentro de su propia carpeta en la raíz del repositorio, como un proyecto independiente. Por ejemplo:

```text
.
├── README.md
├── requirements.txt
├── facturacion/
│   ├── README.md
│   ├── generar_importacion.py
│   ├── files/
│   └── tests/
└── otra_integracion/
    ├── README.md
    ├── src/
    └── tests/
```

El nombre de la carpeta debe describir claramente la integración (por ejemplo, `facturacion`, `inventarios` o `proveedores`). Cada proyecto debe tener su propio `README.md` con su objetivo, alcance, requisitos, configuración, instrucciones de ejecución, pruebas y consideraciones operativas.

## Independencia entre integraciones

- Cada integración debe poder entenderse, configurarse y ejecutarse de manera independiente.
- Mantén dentro de la carpeta del proyecto su código, pruebas, documentación y archivos específicos.
- No mezcles lógica de distintas integraciones. Si en el futuro se comparten componentes, documéntalos y mantenlos en un espacio común claramente identificado.
- No guardes credenciales, contraseñas, tokens ni datos sensibles en el repositorio.
- No sobrescribas archivos de entrada o plantillas originales; genera los resultados en ubicaciones separadas.

## Integración disponible

- [Facturación de compras](facturacion/README.md): prepara el detalle de facturas electrónicas de proveedores para la plantilla de importación de compras de EFFI. El encabezado de la factura permanece manual en el alcance actual.

Consulta el README de cada integración para conocer sus instrucciones específicas. Las dependencias generales de Python se encuentran en [requirements.txt](requirements.txt).
