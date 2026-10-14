# 🛠️ FerreteríaMS AI — Versión 1 (V1 sin IA)

Aplicación web académica desarrollada en **Python 3.12** y **Streamlit** (Ruta A) para consultar productos del inventario de demostración `inventario.xlsx`, verificar existencias (`disponible`, `agotado`, `no encontrado`) y localizar artículos físicamente en tienda mediante coordenadas **`Rack / Columna / Fila`**.

> **Nota importante sobre el alcance de la Versión 1 (V1):**
> Esta primera versión valida la interfaz de usuario, la lectura y validación de datos desde Excel, el motor local de búsqueda (exacta, parcial y ortográfica aproximada) y el soporte bilingüe (español/inglés). **En esta V1 no se integra Gemini, ningún modelo de lenguaje (LLM), ni claves API (`GEMINI_API_KEY`) o secretos (`st.secrets`).** La asesoría inteligente queda pospuesta para una versión futura.

---

## 1. Estructura del proyecto

Los archivos principales en la raíz del repositorio son:

```text
FerreteriaMS_AI/
├── app.py               # Punto de entrada principal de la aplicación Streamlit
├── requirements.txt     # Dependencias del proyecto (streamlit, pandas, openpyxl)
├── inventario.xlsx      # Catálogo de demostración con la hoja 'Inventario'
├── AGENTS.md            # Reglas y directrices de desarrollo del proyecto
├── especificacion.md    # Especificación funcional, contrato de datos y criterios
├── README.md            # Documentación de instalación, uso y despliegue
├── .gitignore           # Exclusión de archivos temporales, entornos y secretos
└── test_app.py          # Pruebas automatizadas de lectura, validación y búsqueda
```

---

## 2. Requisitos e instalación local

- **Python:** versión **3.12** (verificado con Python 3.12.10).
- **Dependencias:** únicamente `streamlit`, `pandas` y `openpyxl`. No requiere claves de API ni conexión a servicios externos de IA.

### Pasos de instalación

1. Clonar o descargar el repositorio y situarse en la carpeta raíz del proyecto:
   ```bash
   cd Ferreteria
   ```
2. (Opcional pero recomendado) Crear y activar un entorno virtual:
   ```bash
   python -m venv .venv
   # En Windows (PowerShell):
   .\.venv\Scripts\Activate.ps1
   # En Linux / macOS:
   source .venv/bin/activate
   ```
3. Instalar las dependencias fijadas en `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```

---

## 3. Ejecución local y pruebas

### Iniciar la aplicación web

Desde la raíz del proyecto, ejecutar:

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en el navegador predeterminado (habitualmente en `http://localhost:8501`).

### Ejecutar las pruebas automatizadas

Para comprobar la sintaxis y ejecutar la batería de pruebas unitarias (lectura de `inventario.xlsx`, validaciones de esquema/stock/ubicación y los casos de búsqueda obligatorios):

```bash
python -m compileall app.py test_app.py
python -m unittest discover -v
```

---

## 4. Estructura y contrato de datos de `inventario.xlsx`

El archivo `inventario.xlsx` debe ubicarse en el mismo directorio que `app.py`. La aplicación resuelve su ruta de manera multiplataforma (Windows/Linux) usando:

```python
INVENTARIO = Path(__file__).resolve().parent / "inventario.xlsx"
```

### Reglas obligatorias de la hoja `Inventario`

- **Nombre exacto de la hoja:** `Inventario`
- **Columnas obligatorias (13 columnas exactas):**
  1. `SKU`: Texto obligatorio, único y sin espacios extremos (ej. `FER-001`).
  2. `Producto`: Texto obligatorio con el nombre del artículo.
  3. `Categoria`: Texto obligatorio para agrupar y filtrar resultados.
  4. `Stock`: Número entero obligatorio $\ge 0$. Si `Stock > 0`, el estado es **Disponible**; si `Stock = 0`, el estado es **Agotado**.
  5. `Rack`: Texto obligatorio cuando `Stock > 0` (ej. `R02`).
  6. `Columna`: Texto obligatorio cuando `Stock > 0` (ej. `A`).
  7. `Fila`: Texto obligatorio cuando `Stock > 0`, conservando ceros iniciales (ej. `03`).
  8. `Descripcion`: Texto opcional con el resumen del catálogo.
  9. `Marca`: Texto opcional.
  10. `Medida`: Texto opcional con la dimensión o presentación (ej. `16 oz`, `1/2 pulgada`).
  11. `Unidad`: Texto opcional (ej. `unidad`).
  12. `Sinonimos`: Texto opcional con términos alternativos en español o inglés separados por `;` (ej. `martillo de bola;hammer;martillo`).
  13. `Especificaciones_tecnicas`: Texto opcional con notas técnicas procedentes exclusivamente del catálogo.

Si el archivo no existe, está dañado, carece de alguna columna obligatoria, tiene algún `SKU` repetido, un `Stock` negativo/no entero o le falta ubicación a un producto con `Stock > 0`, la aplicación detiene la consulta y muestra un mensaje claro de validación sin exponer datos parciales inválidos.

---

## 5. Actualización y recarga del inventario

- **En entorno local:**
  1. Edite y guarde los cambios en la hoja `Inventario` de `inventario.xlsx` desde Excel o LibreOffice Calc, respetando las 13 columnas.
  2. En la interfaz de la aplicación, pulse el botón **«🔄 Recargar inventario»** (o *Reload inventory* en inglés). Esto limpia la caché de lectura (`st.cache_data.clear()`) y vuelve a leer y validar el archivo actualizado.
- **En Streamlit Community Cloud:**
  - El botón **«Recargar inventario»** únicamente vuelve a leer la copia de `inventario.xlsx` que forma parte del despliegue actual; no descarga archivos externos ni ofrece sincronización en tiempo real con sistemas de caja/ERP.
  - Para actualizar los datos del catálogo en la nube, edite y valide `inventario.xlsx` localmente y publique el cambio en el repositorio de GitHub conectado mediante `git commit` y `git push` (previa autorización del propietario).

---

## 6. Limitaciones de la Versión 1 (V1)

- **Sin inteligencia artificial:** No incluye conexión a Gemini, otros LLM, ni chat conversacional. La pestaña de asesoría con IA se muestra explícitamente deshabilitada como referencia del alcance de prueba.
- **Solo lectura:** La aplicación no modifica `inventario.xlsx` ni guarda historiales de búsqueda o datos personales.
- **Sin precios ni recomendaciones inferidas:** Solo se muestran los campos registrados en el Excel. Las especificaciones técnicas se presentan como datos informativos del catálogo, sin inferir compatibilidades entre piezas (por ejemplo, entre `FER-002` y `FER-006`).
- **Privacidad de datos:** Si el repositorio de GitHub es público, todo el contenido de `inventario.xlsx` será visible públicamente. Utilice únicamente datos ficticios de demostración.

---

## 7. Guía paso a paso para GitHub y Streamlit Community Cloud

> **Importante:** No publique el repositorio ni ejecute el despliegue sin la revisión y autorización expresa del responsable del proyecto.

### Paso 1 — Preparar el repositorio en GitHub (con autorización)

1. Verifique que en la raíz del proyecto estén presentes los archivos obligatorios: `app.py`, `requirements.txt`, `inventario.xlsx`, `AGENTS.md`, `especificacion.md`, `README.md` y `.gitignore`.
2. Compruebe que `.gitignore` excluya `.venv/`, `__pycache__/`, `.env` y `.streamlit/secrets.toml`, pero **no excluya `inventario.xlsx`**.
3. Suba los cambios autorizados al repositorio de GitHub:
   ```bash
   git add app.py requirements.txt inventario.xlsx AGENTS.md especificacion.md README.md .gitignore test_app.py
   git commit -m "Preparar FerreteriaMS AI V1 sin IA para Streamlit"
   git push origin main
   ```

### Paso 2 — Desplegar en Streamlit Community Cloud (con autorización)

1. Acceda a [https://share.streamlit.io](https://share.streamlit.io) e inicie sesión conectando su cuenta de GitHub.
2. Pulse el botón **Create app** (o **New app**).
3. Seleccione el **repositorio**, la **rama** (`main`) y especifique `app.py` como **Main file path** (entrypoint).
4. Abra la sección **Advanced settings...**:
   - En **Python version**, seleccione **3.12**.
   - Deje vacío el apartado **Secrets** (esta V1 no utiliza claves de API ni secretos).
5. Pulse **Deploy!** únicamente tras contar con autorización.
6. Una vez completado el despliegue en la URL `*.streamlit.app`, verifique en navegador de escritorio y móvil:
   - Carga inicial sin errores en los logs.
   - Búsquedas por nombre (`Martillo de peña`), medida (`Tornillo 1/2 pulgada`), código (`FER-005`), ortografía aproximada (`martiyo`) y sinónimos en inglés (`hammer`).
   - Funcionamiento del selector de idioma (`Español` / `English`), botón «Limpiar búsqueda» y botón «Recargar inventario».

