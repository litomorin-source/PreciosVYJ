# Historial de Listas de Precios

Primera versión de la aplicación Streamlit para guardar y comparar listas de precios por código.

## Incluye

- Carga de archivos `.xls` y `.xlsx`.
- Detección tentativa de fecha en el título o nombre del archivo.
- Confirmación y edición manual de la fecha antes de guardar.
- Aviso si la fecha ya existe o es anterior a la última carga.
- Guardado del archivo original y de una copia normalizada.
- Historial de cargas.
- Comparación de la última lista contra la anterior.
- Comparación manual entre cualquier par de fechas.
- Mayores y menores aumentos, bajas, altas, retiros y cambios de nombre.
- Buscador por código o producto con evolución histórica.
- Descarga de Excel comparativo.
- Guardado permanente en GitHub mediante Streamlit Secrets.
- Listas iniciales del 07/07/2026 y 16/07/2026.

## Configuración en GitHub y Streamlit

1. Creá un repositorio y subí todo el contenido de esta carpeta.
2. Creá un token Fine-grained de GitHub con acceso únicamente a ese repositorio.
3. En permisos del repositorio habilitá `Contents: Read and write`.
4. En Streamlit Cloud creá una app apuntando a `app.py`.
5. En `Settings > Secrets` cargá:

```toml
GITHUB_TOKEN = "tu_token"
GITHUB_REPO = "usuario/repositorio"
GITHUB_BRANCH = "main"
APP_PIN = "tu_pin"
```

Nunca escribas el token directamente dentro de `app.py`.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```
