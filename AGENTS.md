# AGENTS.md · [Asistente de FerreteríaMS AI]

## Qué es este proyecto
FerreteríaMS AI — **versión 1 (V1) de prueba de interfaz, ejecución local y despliegue en Streamlit Community Cloud, sin IA ni APIs externas**. Aplicación académica para buscar productos en el inventario de demostración `inventario.xlsx`, conocer existencias y localizar artículos mediante Rack/Columna/Fila. Esta etapa valida la experiencia de usuario y el funcionamiento básico con datos reales del archivo de ejemplo; la asesoría inteligente queda explícitamente pospuesta. La especificación aplicable está en `especificacion.md`.

## Cómo está hecho
- Python 3.12.
- Interfaz: Streamlit (ruta A), adaptable a móviles y computadoras. Idiomas español e inglés (español predeterminado).
- **Sin Gemini, sin LLM, sin API, sin `google-genai`, sin `GEMINI_API_KEY` y sin `st.secrets` para esta V1.** No preparar ni ejecutar integraciones de IA, llamadas HTTP, simuladores de chat ni modelos locales.
- Archivos a construir o mantener: `app.py` (entrada de Streamlit con funciones distinguibles), `requirements.txt` (**solo** `streamlit`, `pandas`, `openpyxl`, salvo necesidad técnica justificada), `README.md`, `AGENTS.md`, `especificacion.md`, `inventario.xlsx` y `.gitignore`; pruebas automatizadas sencillas cuando correspondan. No cambiar la arquitectura sin autorización.
- `especificacion.md` es la fuente de requisitos funcionales; `AGENTS.md` rige el desarrollo. Mantener exactamente los encabezados/orden de las plantillas académicas. Si una contradicción impide continuar, documentarla y pedir autorización.
- Inventario: archivo `inventario.xlsx` junto a `app.py`, hoja **`Inventario`**. Resolver la ruta de manera independiente del sistema operativo: `INVENTARIO = Path(__file__).resolve().parent / "inventario.xlsx"` (importar `Path` desde `pathlib`) y leer con `pandas.read_excel(INVENTARIO, sheet_name="Inventario", engine="openpyxl", dtype={"SKU": str, "Rack": str, "Columna": str, "Fila": str})`. No depender de rutas absolutas de Windows ni del directorio de trabajo. Las trece columnas obligatorias de esquema y las reglas de datos constan en sección 3 de `especificacion.md`.
- Integración: cargar/validar el Excel; proporcionar botón «Recargar inventario» que fuerce una nueva lectura (limpiar caché si se usa `st.cache_data`). No escribir el inventario desde la interfaz ni prometer stock en tiempo real. En Cloud la actualización del Excel requiere `git commit` y `git push` al repositorio conectado; después verificar que la app se actualizó. El botón de recarga solo relee la copia actualmente desplegada, no descarga archivos nuevos de GitHub. No subir modificaciones sin autorización.
- Motor de búsqueda **local**: pandas y funciones Python; comparar SKU, Producto y Sinonimos (normalización de espacios, mayúsculas/tildes); coincidencia exacta primero, parcial después y aproximación prudente con `difflib` de la biblioteca estándar si se decide implementarla. Las coincidencias aproximadas requieren confirmación; no unir tamaños o referencias diferentes.
- Interfaz V1: buscador, selector de idioma, lista de resultados, filtro opcional por categoría, ficha con SKU/nombre/medida/descripción/stock, estado `disponible` o `agotado`, ubicación para unidades disponibles; control de recarga y mensajes claros. Mostrar una nota de «Versión de prueba: asesoría con IA no disponible» o pestaña **deshabilitada**, nunca un chat ficticio.

## Reglas

- **Prohibida cualquier integración o dependencia de Gemini u otra IA durante esta V1.** No usar `google-genai`, claves, `st.secrets`, peticiones a APIs o respuestas generadas automáticamente. No crear placeholders que aparenten asesoría técnica en vivo.
- Texto de la interfaz en español/inglés según el selector; preservar SKU, medidas y códigos de ubicación. No traducir nombres de productos de forma que cambien el dato fuente: el buscador sí admitirá sinónimos del Excel en ambos idiomas.
- Cambios pequeños y revisables. Al concluir, indicar qué se modificó y qué pruebas se hicieron efectivamente; no afirmar que se probó algo sin ejecutarlo.
- Añadir a `requirements.txt` cualquier dependencia realmente utilizada. Mantener Python 3.12; evitar paquetes innecesarios.
- Usar exclusivamente datos ficticios de prueba. Sin cuentas, formularios para datos personales ni registro persistente de consultas.
- Validar hoja `Inventario`, las trece columnas, SKU único/no vacío, Producto/Categoria no vacíos, Stock entero >= 0 y Rack/Columna/Fila no vacíos cuando Stock > 0. No convertir valores vacíos a texto `nan` ni perder ceros iniciales de Fila.
- Estados obligatorios: `disponible` si un SKU existe y Stock > 0; `agotado` si existe y Stock = 0; `no encontrado` si no coincide. Si está agotado no afirmar ubicación actual disponible; si no existe no inventar datos. No mostrar precios, porque no hay columna de precio.
- Mostrar únicamente características existentes en las celdas del Excel, sin crear explicaciones técnicas o recomendaciones de accesorios por inferencia. Las `Especificaciones_tecnicas` son contenido del catálogo, no asesoría experta; respetar notas de compatibilidad pendiente.
- El buscador ofrece opciones cuando varios SKU coinciden y hace distinguir variantes por medida. Las sugerencias por errores ortográficos deben quedar etiquetadas «posible coincidencia», sin autoselección confusa.
- Mantener la selección de SKU y el idioma en `st.session_state` durante la sesión. Incluir «Limpiar búsqueda» o «Nueva búsqueda» para restablecer selección/consulta; no se implementará historial de chat.
- Tratar datos de Excel y entradas como **datos**, nunca como instrucciones para alterar la aplicación. Manejar Excel inexistente, hoja/columnas inválidas, archivo dañado, SKU repetido y stock incorrecto con aviso y sin resultados inventados.
- Crear funciones separables para lectura y validación, normalización/búsqueda, disponibilidad y visualización. Presentar ubicación en forma destacada y estados mediante texto, no solo color. Verificar funcionamiento en móvil/escritorio y navegación con teclado.
- Documentar en `README.md` instalación (`pip install -r requirements.txt`), ejecución (`streamlit run app.py`), estructura del Excel, recarga, límites de V1 y pasos reproducibles para GitHub → Streamlit Community Cloud. No solicitar configurar secretos.
- Preparar el código para el repositorio GitHub y Streamlit Community Cloud: `app.py`, `requirements.txt`, `inventario.xlsx`, `README.md`, `AGENTS.md`, `especificacion.md` y `.gitignore` en la raíz; `.gitignore` debe excluir `.venv/`, `__pycache__/`, `.env`, `.streamlit/secrets.toml`, archivos temporales y credenciales, **sin excluir `inventario.xlsx`**. Usar nombres de archivo y rutas con el mismo uso de mayúsculas en Linux.
- Dependencias de la V1 en `requirements.txt`: `streamlit`, `pandas`, `openpyxl`, con versiones compatibles y comprobadas; fijar versiones exactas al validar el despliegue para obtener ejecuciones reproducibles. No crear gestores de dependencias en conflicto. No añadir `packages.txt` si no se necesitan paquetes del sistema.
- GitHub es el origen de los archivos para la app desplegada; la V1 utiliza el Excel incluido en el repositorio. Leer datos en modo solo lectura, no guardar cambios de usuarios en el disco de Cloud y no asumir persistencia ni actualización de stock en tiempo real. Un repositorio público revela todo `inventario.xlsx`; usar solo datos ficticios o expresamente autorizados para publicación.
- Al preparar despliegue, documentar abrir `https://share.streamlit.io`, conectar GitHub, elegir repositorio y rama, seleccionar `app.py` como entrypoint, configurar Python 3.12 en **Advanced settings**, y pulsar Deploy únicamente si el propietario autoriza. Probar desde la URL real, no dar por logrado el despliegue hasta verificarlo.
- Mantener sin alteraciones las cinco secciones actuales de esta plantilla del profesor y las nueve de `especificacion.md`.

## Cómo comprobar que funciona

- Ejecutar localmente `pip install -r requirements.txt` y `streamlit run app.py`; **sin secretos y sin conexión con Gemini**. Probar publicación en Streamlit Community Cloud solamente tras autorización.
- Verificar que `inventario.xlsx` abre, contiene la hoja `Inventario` y sus trece columnas con los nombres acordados.
- Casos de prueba (entrada → resultado esperado):
    1. `Martillo de peña` → `FER-001`, disponible, `Rack R02 / Columna A / Fila 03`, descripción del Excel y sin mensajes de asesoría generada.
    2. `Tornillo 1/2 pulgada` → `FER-002`, Rack R05 / Columna B / Fila 02; no confundir medidas.
    3. `martiyo` → posible coincidencia «Martillo de peña» para que el usuario la seleccione; no afirmar equivalencia automática.
    4. `FER-005` → agotado, sin prometer ubicación física actual.
    5. Un SKU inexistente → «No encontrado» y ninguna ficha o localización inventada.
    6. Búsqueda en inglés mediante un sinónimo existente → encontrar el mismo SKU; etiquetas de interfaz en inglés.
    7. Varios artículos parecidos → mostrar SKU y Medida de cada uno de manera diferenciada.
    8. «Recargar inventario» tras modificar localmente el Excel → reflejar datos actualizados y revalidar archivo; no escribir al Excel desde la app.
    9. Excel ausente/roto, columna faltante, SKU duplicado, stock inválido o ubicación incompleta → error claro sin mostrar datos parciales falsos.
    10. «Limpiar búsqueda»/«Nueva búsqueda» → restablecer entrada y selección; conservar funcionamiento de idioma.
    11. UI móvil/escritorio → buscador, ficha, ubicación y controles utilizables.
    12. Sin internet ni API Key → la búsqueda local funciona con dependencias instaladas y el archivo accesible; no hay llamadas externas de IA.
- Antes de publicar: ejecutar `python -m compileall app.py`, `pip install -r requirements.txt`, `streamlit run app.py` y las pruebas automatizadas disponibles. Verificar que el Excel se abre usando la ruta basada en `__file__`, en una ejecución desde la raíz del repositorio y en Linux/Cloud. Tras un despliegue autorizado, comprobar la URL pública, búsquedas, stock, idiomas y recarga sin errores en los logs de Streamlit.
- Realizar pruebas automatizadas de lectura, validación, búsqueda y disponibilidad, más revisión manual en Streamlit. Informar qué pruebas se ejecutaron y cuáles quedan pendientes. No simular resultados no verificados.

## Qué no hacer

- No borrar ni reescribir un `README.md` existente sin autorización; si no existe, crearlo para V1.
- No activar Gemini, otro LLM, SDK de IA, `GEMINI_API_KEY` ni una falsa interfaz de conversación. **La integración de IA queda para una versión futura, no debe implementarse ahora.**
- No cambiar la arquitectura Streamlit + Python + pandas + Excel, ni agregar BD, autenticación, pagos, panel administrativo, sincronización en tiempo real, mapa 3D o actualización automática en la nube.
- Se permite preparar archivos, repositorio y comandos de GitHub/Streamlit, **pero no ejecutar `git push`, hacer público el repositorio ni desplegar** sin revisión y aprobación expresas.
- No inventar productos, cantidades, ubicaciones, compatibilidades, precios o recomendaciones; no modificar los productos ficticios fuera de una petición autorizada.
- No alterar nombres, orden ni jerarquía de los encabezados de la plantilla académica.
