# CATIA V5 Data Extractor & Snapshot Tool

Este proyecto es una herramienta automatizada desarrollada en **VB.NET** diseñada para extraer estructuras de productos de **CATIA V5** y generar reportes técnicos detallados en **Excel**, incluyendo capturas de pantalla automáticas de cada componente.

## Funcionalidades Clave
* **Extracción de Estructura:** Recorre ensamblajes complejos de forma recursiva.
* **Generación de Snapshots:** Captura automáticamente imágenes en vista isométrica con fondo normalizado (blanco) para cada pieza.
* **Consolidación de Datos:** Agrupa componentes por *Part Number* y calcula cantidades automáticamente.
* **Filtro Inteligente de Componentes:** Detecta y omite "Internal Components" (nodos sin archivo físico) para mantener un reporte limpio, procesando directamente sus hijos.
* **Reporte Profesional:** Genera un archivo Excel formateado con niveles jerárquicos y metadatos técnicos.



## Requisitos Técnicos
* **Lenguaje:** VB.NET (.NET Framework)
* **Software:** CATIA V5 (instalado y con licencia activa).
* **Referencias COM:** * `INFITF`
    * `ProductStructureTypeLib`
    * `Microsoft.Office.Interop.Excel`

## Estructura del Código
* `CatiaDataExtractor`: Motor de recursividad y detección de jerarquías.
* `TakeSnapshot`: Módulo de gestión visual y control de ventanas de CATIA.
* `ExcelFormatter`: Encargado de la estética y estructura del reporte final.
* `PwrProduct`: Clase de datos para el almacenamiento temporal de propiedades.

---
Generado para optimizar procesos de ingeniería y automatización de diseño 3D.
