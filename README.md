# Informe de cobertura y reposición — versión corregida

Se corrigió el reconocimiento de extensiones Excel `.xlsx` y `.xls`, el análisis de fechas de período y la lectura del reporte de consumo Excel.

- El consumo Excel se lee desde la primera hoja.
- Se buscan encabezados Código, Cantidad y Tipo Documento entre las primeras 60 filas y todas las columnas.
- Si no se reconocen por nombre, se aplica respaldo de columnas E, H y J, cuando esas celdas de encabezado identifican Código, Cantidad y Tipo Documento.
- El resto de las funciones de la aplicación se mantiene.

Para publicar: subir el contenido de esta carpeta a la raíz del repositorio GitHub y activar GitHub Pages desde Settings → Pages.
