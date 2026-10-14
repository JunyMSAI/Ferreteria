# Especificación · [Asistente de FerreteríaMS AI]

## 1. Problema y usuario

- **Quién tiene el problema:**
Clientes de ferretería que pierden tiempo buscando artículos en racks/estantes y necesitan saber si se encuentran disponibles.

- **Cómo lo resuelve hoy:**
Esperan a que un empleado los ayude a localizar productos y verificar existencias.

- **Qué le molesta de cómo lo hace hoy:** [tiempo, errores, no encuentra la información…]
Tiempo de espera, desconocimiento de ubicaciones y falta de claridad sobre las medidas o disponibilidad.

- **Qué parte resuelve la IA:** [y qué parte se queda fuera]
**En esta V1 no se utiliza IA.** Python consulta directamente el Excel para localizar productos y mostrar datos ya registrados; no genera asesoría técnica ni identifica accesorios automáticamente. La IA y la conversación contextual son una ampliación futura, no un requisito ni una funcionalidad simulada en esta prueba.

## 2. Qué hace el producto

Máximo tres funcionalidades. Si hay más, van al apartado 8.

1. Buscar artículos del archivo `inventario.xlsx` por SKU, nombre o sinónimos, distinguiendo coincidencias exactas, parciales y sugerencias aproximadas.
2. Mostrar ficha del producto (nombre, categoría, SKU, medida, descripción y datos técnicos **tal como constan en el Excel**) y su estado: disponible o agotado; si hay stock, mostrar Rack/Columna/Fila.
3. Permitir probar la experiencia de usuario en español e inglés: elegir entre resultados, navegar por la ficha, limpiar la búsqueda y recargar el Excel; ofrecer diseño adaptable a móvil y escritorio. **Sin asesoría de IA ni chat en V1.**

## 3. Entradas y salidas

| El usuario da… | La app devuelve… | Ejemplo |
|---|---|---|
| Nombre, SKU o sinónimo de producto | Coincidencias del Excel identificadas por SKU y medida, estado de stock, ubicación Rack/Columna/Fila si disponible y ficha de producto | «Martillo de peña» → `FER-001`, disponible, Rack R02 / Columna A / Fila 03 |
| Descripción de una tarea | No se ofrece asesoría en esta versión; puede buscar artículos concretos por su nombre o SKU | «Necesito sujetar una tubería» → mensaje de que la asesoría no está disponible; usar búsqueda |
| Producto con Stock = 0 | Estado «agotado» sin promesa de disponibilidad o ubicación física actual | `FER-005` → Agotado |
| Producto que no aparece | Mensaje «no encontrado», sin guía generada ni ubicación inventada | «Tornillo de 3 pulgadas» → No encontrado en el catálogo de ejemplo |
| Consulta fuera de alcance | Respuesta breve delimitando temática de ferretería | «¿Cuál es la capital de Francia?» → Solo puedo ayudarte con temas de ferretería y bricolaje |

**Contrato de datos para `inventario.xlsx`:** archivo situado junto a `app.py`, con una hoja cuyo nombre exacto es `Inventario`. La primera fila es el encabezado; una fila por SKU. Se deben utilizar exactamente estas trece columnas, respetando mayúsculas y tildes: `SKU`, `Producto`, `Categoria`, `Stock`, `Rack`, `Columna`, `Fila`, `Descripcion`, `Marca`, `Medida`, `Unidad`, `Sinonimos`, `Especificaciones_tecnicas`.

| Columna | Tipo y regla | Finalidad |
|---|---|---|
| `SKU` | Texto obligatorio, único, sin espacios extremos | ID estable del producto; conservar ceros iniciales |
| `Producto` | Texto obligatorio | Nombre mostrado al cliente |
| `Categoria` | Texto obligatorio | Agrupar y filtrar resultados |
| `Stock` | Entero obligatorio, >= 0 | `>0` disponible; `0` agotado; no permitir negativos |
| `Rack` | Texto; obligatorio si Stock > 0 | Código de rack, ej. R02 |
| `Columna` | Texto; obligatorio si Stock > 0 | Columna, ej. A |
| `Fila` | Texto; obligatorio si Stock > 0 | Fila, ej. 03; conservar ceros iniciales |
| `Descripcion` | Texto opcional | Información resumida basada en datos del producto |
| `Marca` | Texto opcional | Marca declarada; no inventarla |
| `Medida` | Texto opcional | Diámetro, longitud u otra medida tal como figure |
| `Unidad` | Texto opcional | Unidad de venta, ej. unidad, paquete |
| `Sinonimos` | Texto opcional; entradas separadas por `;` | Búsqueda alternativa, ej. martillo de bola; hammer |
| `Especificaciones_tecnicas` | Texto opcional | Material, tipo de rosca, estándar o compatibilidad solo cuando estén comprobados |

**Integración Excel → Python → Streamlit:** instalar `pandas` y `openpyxl` mediante `requirements.txt`; colocar `inventario.xlsx` en la raíz del proyecto; resolver ruta con `Path(__file__).resolve().parent / "inventario.xlsx"` y cargar con `pd.read_excel(ruta_inventario, sheet_name="Inventario", engine="openpyxl", dtype={"SKU": str, "Rack": str, "Columna": str, "Fila": str})`. Esta ruta relativa al archivo de aplicación funciona en Windows y Linux/Cloud. Validar presencia de hoja, columnas, SKU, stock y ubicación antes de buscar. Tratar celdas vacías como valores ausentes, no como texto `nan`; normalizar solo una copia destinada a búsqueda (acentos, espacios y capitalización) sin modificar los valores originales mostrados. La app consulta directamente el DataFrame, sin usar LLM, red o API para inventario.

**Actualización del inventario:** el administrador modifica la hoja `Inventario` fuera de la app, manteniendo las columnas y el SKU de cada referencia. En local se vuelve a cargar el archivo al reiniciar o pulsar «Recargar inventario» (invalidar `st.cache_data` si se usa). En Streamlit Community Cloud el Excel se publica como archivo del repositorio, mediante commit/push autorizado y posterior actualización de la app; las modificaciones hechas solo al disco efímero de la app desplegada **no son persistentes**. No hay sincronización automática con caja/ERP ni confirmación en tiempo real de inventario en esta fase.

**Comportamiento de búsqueda:** buscar por SKU, nombre y sinónimos, con coincidencia exacta preferida, luego parcial y finalmente sugerencia aproximada sin selección automática si hay ambigüedad. Manejar ortografía sencilla como «martiyo», sin asumir compatibilidad entre tamaños o artículos parecidos. Mostrar cada variante/SKU por separado. Estados inequívocos: `disponible`, `agotado`, `no encontrado`. Si Stock > 0, mostrar los tres componentes de localización; si Stock = 0, mostrar estado agotado sin indicar ubicación actual confirmada. Si no figura, no fabricar referencias ni datos.

## 4. La pieza de IA

- **Opción elegida:**
1. **V1 SIN IA:** ninguna API, modelo de IA, servicio de Gemini, SDK `google-genai`, credencial ni mecanismo de chat. El título de esta sección se conserva por ser parte de la plantilla del profesor.
2. Python 3.12 → librerías `pandas` y `openpyxl` para Excel, `streamlit` para UI; `difflib` (biblioteca estándar de Python) opcional para sugerencias ortográficas.

- **Por qué:**
1. Esta primera prueba sirve para evaluar interfaz, navegación, búsqueda, accesibilidad, idioma, lectura del inventario y precisión de ubicación **sin depender de un LLM ni incurrir en llamadas externas**.
2. pandas permite cargar y buscar en un Excel estructurado por SKU/Rack/Columna/Fila; los datos provienen exclusivamente del archivo local, evitando ubicaciones y stock inventados.

- **Instrucciones para el modelo (system prompt):**
**No aplica en la V1.** No redactar ni utilizar system prompts, instrucciones de rol para IA, peticiones a Gemini o imitaciones de respuestas inteligentes. En la interfaz, si existe una sección visual reservada para futura asesoría, mostrar únicamente «Asesoría con IA: no disponible en esta versión de prueba» (en inglés cuando corresponda) y mantenerla deshabilitada. No captar mensajes para un chat inexistente.

**Comportamiento local en sustitución de IA:** la aplicación solo puede buscar, filtrar y mostrar las trece columnas de `Inventario`, con texto fijo traducido para controles y mensajes. Puede mostrar la `Descripcion` y `Especificaciones_tecnicas` originales; no sintetizar diagnósticos ni instrucciones de seguridad nuevas. La selección actual y el idioma pueden mantenerse temporalmente con `st.session_state`, sin historial conversacional. Los requisitos de IA quedan fuera de alcance hasta una solicitud posterior.

## 5. Flujo del usuario

1. El cliente abre Streamlit en móvil o escritorio y ve una indicación visible «Versión de prueba sin IA», buscador y selector de idioma (español predeterminado).
2. Introduce nombre, SKU o sinónimo; la app compara con datos de `inventario.xlsx` y lista SKU/medidas distintos. Las coincidencias aproximadas aparecen como posibles opciones, no como selecciones automáticas.
3. Selecciona un producto; ve nombre, SKU, categoría, medida, descripción, existencias y estado. Si Stock > 0, muestra `Rack / Columna / Fila`; con Stock = 0, informa «Agotado» y no asegura ubicación física actual. Los datos técnicos se presentan **solo si aparecen en Excel**, sin interpretaciones de IA.
4. El usuario puede limpiar la búsqueda, elegir otro producto, cambiar idioma o pulsar «Recargar inventario». Si se visualiza una sección futura de asesoría, se muestra claramente deshabilitada; **no hay mensajes/chat ni recomendaciones automáticas**.

**Presentación esperada:** búsqueda siempre visible, resultados con SKU/medida, estados visibles por texto además del color, ubicación destacada, mensajes de validación comprensibles, controles accesibles en móvil y por teclado. Etiquetas y mensajes predefinidos en español/inglés; SKU y datos originales del archivo conservados.

## 6. Restricciones

- V1 de prueba sin Gemini, modelos LLM, API Key, `st.secrets`, llamadas externas de IA ni facturación. La interfaz se ejecuta localmente y está preparada para desplegarse en Streamlit Community Cloud desde GitHub (sujeto a límites del servicio y autorización de publicación).
- Solo Python 3.12, Streamlit, pandas y openpyxl (más biblioteca estándar); no instalar `google-genai` ni crear configuración de secretos.
- Sin datos personales reales ni persistencia de consultas. El Excel incluye datos de demostración ficticios.
- Mostrar información exclusivamente del inventario, sin recomendaciones de accesorios, instrucciones técnicas inventadas o promesas de compatibilidad. No presentar la app como asistente inteligente activo en V1.
- Textos de interfaz en español/inglés con palabras comprensibles; errores sin trazas completas para el usuario.
- Validar: archivo inexistente/dañado, hoja o columnas faltantes, SKU duplicado, Stock negativo/no entero y ubicación incompleta con existencias positivas; bloquear resultados inválidos y mostrar un error claro.
- No hay edición del inventario desde la app, sincronización en tiempo real, pagos, usuarios ni panel administrativo. El Excel se administra fuera de Streamlit.
- No publicar un catálogo real o confidencial en un repositorio público. No desplegar sin autorización.

## 7. Criterios de aceptación

La app está bien si:

| Entrada | Resultado esperado |
|---|---|
| «Martillo de peña» | Busca en Excel y muestra `FER-001`, disponible, Rack R02 / Columna A / Fila 03; sin guía generada por IA |
| «martiyo» | Sugiere «Martillo de peña» como posible coincidencia sin sustituir automáticamente el SKU |
| «Tornillo 1/2 pulgada» | Muestra `FER-002`, sus datos y Rack R05 / Columna B / Fila 02; no cambia la medida |
| `FER-005` | Detecta Stock = 0 e informa «agotado» sin prometer existencia en el rack |
| Producto no registrado | «No encontrado», sin SKU o ubicación inventada |
| Productos similares | Lista opciones con SKU y medidas, y permite seleccionar una ficha concreta |
| Idioma inglés | Cambia etiquetas, controles y mensajes a inglés; mantiene los datos/SKU del Excel |
| «Recargar inventario» | Refresca lectura de archivo y aplica validaciones de nuevo |
| Excel ausente/dañado, columnas incorrectas, SKU duplicado o Stock inválido | Mensaje claro; no muestra información parcial no validada |
| «Limpiar búsqueda» / «Nueva búsqueda» | Limpia consulta y selección sin recargar obligatoriamente la página |
| Sin internet, sin clave de API | Funciona la consulta local (con dependencias ya instaladas) sin llamar a Gemini ni a servicios de IA |
| Zona de IA, si se muestra | Está inactiva y dice «No disponible en esta versión de prueba»; no ofrece chat |
| Vista móvil | Todos los controles y ubicaciones son legibles y utilizables sin desbordamiento horizontal obligatorio |

**Pruebas exigidas al desarrollador (Gemini como herramienta para *escribir código*, no como API dentro de la app):** ejecutar pruebas locales de lectura, validación, normalización, búsqueda y casos negativos y revisar manualmente la interfaz en ambos idiomas y tamaños. Reportar resultados reales, sin afirmar pruebas no ejecutadas.

## 8. Fuera de alcance (por ahora)

- Mapa 3D interactivo de la ferretería.
- **Integración de Gemini o cualquier IA, chat conversacional, asesoría técnica generada y recomendaciones de accesorios automáticas.** Se dejan para una fase posterior, sin código ni dependencias de IA en esta V1.
- POS/ERP en tiempo real, códigos de barras, pagos, cuentas de clientes, panel administrativo, edición del Excel desde la app y sincronización automática con proveedores.

## 9. Despliegue

- **Ruta:**
Ruta A: aplicación Python 3.12 + Streamlit, desarrollada primero localmente y **preparada para ejecutarse desde GitHub mediante Streamlit Community Cloud**, después de autorización expresa. Gemini CLI puede crear los archivos y ejecutar pruebas locales, pero no publicar ni desplegar sin permiso.

**Estructura obligatoria en la raíz del repositorio:**

```text
FerreteriaMS_AI/
├── app.py
├── requirements.txt
├── inventario.xlsx
├── AGENTS.md
├── especificacion.md
├── README.md
└── .gitignore
```

**Archivos y configuración:** `app.py` es el entrypoint. `requirements.txt` incluye solo `streamlit`, `pandas` y `openpyxl` con versiones compatibles comprobadas (preferiblemente fijadas tras probar). No se requiere `GEMINI_API_KEY`, `google-genai`, `st.secrets` ni secretos para esta V1. `.gitignore` excluye entornos virtuales, cachés, `.env`, `.streamlit/secrets.toml`, credenciales y archivos temporales, pero **no debe excluir** `inventario.xlsx`. No incluir rutas absolutas ni separadores `\` de Windows al cargar archivos: usar `pathlib.Path(__file__).resolve().parent / "inventario.xlsx"`.

**Preparación en GitHub (con aprobación del propietario):** (1) verificar que únicamente se publican datos ficticios y no existen secretos; (2) crear o seleccionar repositorio GitHub; (3) subir los siete archivos de la estructura, mantener exactamente los nombres y confirmar que `inventario.xlsx` está versionado; (4) mantener en `README.md` instrucciones locales, dependencias, datos de demostración y procedimiento de actualización; (5) no hacer público ni ejecutar `git push` sin revisión y autorización expresa. Repositorios privados pueden requerir permisos adicionales de conexión con Streamlit.

**Despliegue en Streamlit Community Cloud (con autorización):** entrar a `https://share.streamlit.io`, conectar la cuenta de GitHub, seleccionar **Create app** y el repositorio/rama, indicar `app.py` como archivo de entrada, y en **Advanced settings** escoger Python **3.12**. No introducir secretos: la app no usa APIs. Desplegar y comprobar la URL `*.streamlit.app`, el arranque sin errores, la carga del Excel, búsqueda, estados, idioma y diseño móvil. El alojamiento gratuito y sus límites dependen de Streamlit.

**Actualizaciones:** editar y validar el Excel localmente y publicar la nueva versión en GitHub mediante commit/push autorizado; Streamlit Cloud normalmente actualiza la app conectada al repositorio. La opción «Recargar inventario» solo vuelve a leer el Excel de **esa versión ya desplegada**; no descarga cambios externos ni proporciona stock en tiempo real. El sistema de archivos de Cloud no se usa como almacenamiento persistente para modificaciones de inventario.

**Pruebas de aceptación del despliegue:** instalación sin dependencias ausentes; Python 3.12; `app.py` identificable; Excel incluido y legible en entorno Linux; datos coherentes con la hoja; ausencia de secretos y llamadas a IA; resultados correctos y accesibles desde la URL de Cloud. Si falla un despliegue, revisar logs y corregir el repositorio; nunca afirmar que Cloud fue probado cuando solo se hicieron pruebas locales.
