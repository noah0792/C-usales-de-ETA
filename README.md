# Consulta de Causales HFC 2026

Aplicación web para consultar códigos de causales (HFC, CANOPY, GPON, Hoja2).
Se escribe un código o una palabra, se selecciona de la lista y aparece su
descripción completa, con su ubicación, tipo y causa, tal como está en la tabla.

## Archivos

- `index.html`: la aplicación. Funciona sola, sin servidor ni instalaciones.
- `generar_datos.py`: lee el Excel y regenera `index.html`.

## Cómo regenerar la aplicación

Si cambia el Excel, ejecute:

```bash
pip install openpyxl
python generar_datos.py TABLA_CAUSALES_HFC_2026.xlsx index.html
```

## Publicar en GitHub Pages

1. Cree un repositorio nuevo en GitHub (por ejemplo `causales-hfc-2026`).
2. Suba `index.html`, `generar_datos.py` y este `README.md`.
   Opcional: suba también el Excel para poder regenerar.
3. En el repositorio vaya a **Settings → Pages**.
4. En **Build and deployment**, elija **Deploy from a branch**, rama `main` y carpeta `/ (root)`.
5. Guarde. En uno o dos minutos la app estará en `https://SU-USUARIO.github.io/causales-hfc-2026/`.

## Notas sobre los datos

- Los textos se muestran exactamente como están en el Excel, incluidos sus errores de escritura.
- Un mismo código puede aparecer en varias secciones con distinta descripción (por ejemplo, 8427 en HFC).
  La aplicación muestra todas las coincidencias, y cada una indica su sección y producto.
- Las celdas con fórmulas (por ejemplo `=+J5`) se muestran con el valor que tienen calculado en el Excel.
- Las hojas CANOPY y Hoja2 no tienen encabezados en el Excel, así que sus etiquetas son genéricas.
