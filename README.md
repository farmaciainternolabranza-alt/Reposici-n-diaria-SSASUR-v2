# Aplicación de reposición – Farmacia CESFAM Labranza

Aplicación estática para GitHub Pages. Lee el CSV de consumo y el XLSX de stock directamente en el navegador.

## Publicar en GitHub Pages

1. Descarga y descomprime el ZIP.
2. Crea un repositorio nuevo en GitHub.
3. Sube **el contenido** de esta carpeta (index.html, styles.css, app.js y README.md) a la raíz del repositorio.
4. En GitHub abre **Settings → Pages**.
5. En “Build and deployment”, selecciona **Deploy from a branch**.
6. Selecciona la rama `main` y la carpeta `/ (root)`, y guarda.
7. Espera la publicación. GitHub mostrará la dirección de la aplicación en esa misma sección.

No se requiere instalar nada en el computador. Para leer XLSX se carga SheetJS desde jsDelivr, por lo que se necesita conexión a internet al abrir la aplicación. Los archivos de trabajo no se envían a ese servicio: la librería se descarga y el procesamiento ocurre localmente en el navegador.

## Uso

1. Abre la dirección publicada.
2. Arrastra el reporte de consumo `.csv` al recuadro de consumo.
3. Arrastra el reporte de stock `.xlsx` al recuadro de stock.
4. Selecciona los tipos de documento que deben formar el consumo.
5. Asigna una bodega a “Stock Farmacia” y otra distinta a “Stock Bodega”. Las demás se excluyen.
6. Pulsa “Procesar y generar informe”.
7. Imprime o descarga el resultado como CSV.

## Reglas implementadas

- Consumo: detecta la fila de encabezados por `Código`, `Cantidad` y `Tipo Documento`; agrupa por código y suma cantidad solo para los tipos seleccionados.
- Los tipos de documento se detectan automáticamente y aparecen todos seleccionados inicialmente.
- Stock: lee la primera hoja del XLSX y detecta la fila por `Código`, `Bodega` y `Cantidad`.
- Detecta la fecha `Obtenido : dd/mm/aaaa` en las filas anteriores a los encabezados. Solo permite procesar si coincide con la fecha actual de `America/Santiago`.
- Reconoce las bodegas existentes; exige seleccionar dos distintas. Todas las demás quedan fuera.
- Cruza consumo y stock por código.
- Prioridad 1: stock Farmacia igual a cero y stock Bodega mayor que cero.
- Luego ordena por disponibilidad ascendente: stock Farmacia / consumo × 100.
- Si consumo es cero, muestra `N/C` y ordena esos artículos al final.
- El período final del consumo se compara con el último período procesado y se guarda en `localStorage`. Si retrocede, muestra advertencia, pero permite continuar.
- Incluye impresión y descarga CSV.

## Consideraciones y límites

- El CSV debe contener encabezados reconocibles: `Código`, `Cantidad`, `Tipo Documento`; el separador puede ser coma, punto y coma o tabulación.
- El período de consumo debe aparecer en el texto como `desde dd/mm/aaaa hasta dd/mm/aaaa`.
- El XLSX debe tener en la primera hoja encabezados `Código`, `Bodega`, `Cantidad` y una fecha `Obtenido : dd/mm/aaaa` en las filas superiores.
- Se detectan columnas por el texto de sus encabezados, no por una letra fija. Esto evita depender de desplazamientos de columnas entre reportes.
- La aplicación no inventa nombres de artículos: el informe identifica cada registro por código. Si el XLSX incluye una columna `Descripción`, `Producto` o `Artículo`, se muestra como nombre del artículo; el cruce sigue realizándose exclusivamente por código.
- El primer informe procesado establece la fecha de referencia para comparar períodos posteriores. El dato se guarda en el navegador y perfil donde se usa la aplicación; borrar los datos del sitio o cambiar de navegador/perfil reinicia esa referencia.
- Los archivos no se guardan ni transmiten. La aplicación no tiene backend.
- La librería SheetJS se obtiene desde CDN; por tanto, se requiere conexión a internet para cargarla.

## Pruebas recomendadas antes de uso operativo

Probar con copias de reportes reales:
1. CSV con varios tipos de documento y códigos repetidos.
2. XLSX de hoy con tres o más bodegas.
3. Stock con fecha de ayer (debe bloquear).
4. Consumo cuyo período final retrocede (debe advertir y permitir continuar).
5. Artículo con stock Farmacia cero y stock Bodega positivo (debe aparecer primero).
6. Artículo con consumo cero (debe mostrar N/C).
