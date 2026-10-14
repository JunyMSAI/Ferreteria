"""FerreteríaMS AI · Versión 1 (V1) de prueba de interfaz sin IA.

Aplicación web en Streamlit (Ruta A) para consultar productos del archivo local
`inventario.xlsx`, conocer existencias (`disponible`, `agotado`, `no encontrado`)
y localizar artículos mediante `Rack / Columna / Fila`.

En esta primera versión NO se integra Gemini, ningún LLM, ni APIs externas.
"""

from __future__ import annotations

import difflib
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

# =====================================================================
# 1. CONSTANTES Y CONFIGURACIÓN DE INVENTARIO
# =====================================================================

INVENTARIO: Path = Path(__file__).resolve().parent / "inventario.xlsx"
HOJA_INVENTARIO: str = "Inventario"

COLUMNAS_OBLIGATORIAS: list[str] = [
    "SKU",
    "Producto",
    "Categoria",
    "Stock",
    "Rack",
    "Columna",
    "Fila",
    "Descripcion",
    "Marca",
    "Medida",
    "Unidad",
    "Sinonimos",
    "Especificaciones_tecnicas",
]

TITULO = "🛠️ FerreteríaMS AI"
DESCRIPCION = (
    "Buscador de catálogo, existencias y ubicación física en tienda (Rack / Columna / Fila)."
)

# Diccionario de internacionalización (Español predeterminado / Inglés)
I18N: dict[str, dict[str, str]] = {
    "es": {
        "lang_label": "🌐 Idioma / Language",
        "app_title": "🛠️ FerreteríaMS AI — Localizador de Productos",
        "app_subtitle": "Consulta rápida de existencias y ubicación física en tienda a partir del catálogo oficial.",
        "v1_badge": "ℹ️ Versión de prueba (V1 sin IA): Asesoría con IA no disponible en esta versión de prueba. La búsqueda consulta directamente el archivo inventario.xlsx.",
        "reload_btn": "🔄 Recargar inventario",
        "reload_help": "Vuelve a leer el archivo inventario.xlsx local y actualiza la validación de datos.",
        "reload_ok": "Inventario recargado y validado correctamente.",
        "tab_search": "🔍 Buscador de Catálogo",
        "tab_ai_disabled": "🤖 Asesoría con IA (No disponible en V1)",
        "ai_disabled_title": "Asesoría con IA: no disponible en esta versión de prueba",
        "ai_disabled_desc": (
            "Esta primera versión (V1) valida la interfaz, la búsqueda por SKU/nombre/sinónimos, "
            "el control de existencias y la localización por Rack / Columna / Fila sin conexión a modelos "
            "de inteligencia artificial ni chats simulados. Para encontrar un artículo, utilice la pestaña "
            "«Buscador de Catálogo»."
        ),
        "search_label": "Buscar por nombre de producto, código SKU o sinónimo",
        "search_placeholder": "Ej.: Martillo de peña, FER-001, Tornillo 1/2 pulgada, hammer...",
        "category_label": "Filtrar por categoría (opcional)",
        "all_categories": "Todas las categorías",
        "search_btn": "🔎 Buscar",
        "clear_btn": "🧹 Limpiar búsqueda / Nueva búsqueda",
        "results_header": "Resultados de la búsqueda",
        "catalog_header": "Catálogo de productos disponibles para consulta",
        "exact_matches_msg": "Se encontraron {n} coincidencia(s) exacta(s) o directa(s):",
        "partial_matches_msg": "Se encontraron {n} coincidencia(s) en el catálogo:",
        "approx_warning_title": "⚠️ Posible coincidencia (sugerencia por ortografía aproximada)",
        "approx_warning_desc": (
            "No se encontró una coincidencia exacta para «{q}». "
            "Se muestran posibles coincidencias aproximadas. Confirme y seleccione manualmente el artículo "
            "deseado verificando su SKU y medida:"
        ),
        "select_product_prompt": "Seleccione un artículo de la lista para ver su ficha completa y ubicación:",
        "view_sheet_btn": "Ver ficha de {sku}",
        "selected_btn": "✓ Viendo ficha de {sku}",
        "status_available": "DISPONIBLE",
        "status_out_of_stock": "AGOTADO",
        "status_not_found": "NO ENCONTRADO",
        "not_found_title": "❌ Producto no encontrado",
        "not_found_desc": (
            "No se encontró ningún artículo que coincida con «{q}» en el catálogo de demostración. "
            "Verifique el código SKU, el nombre o intente con términos más simples."
        ),
        "task_query_title": "ℹ️ Asesoría técnica no disponible en esta versión de prueba",
        "task_query_desc": (
            "En esta versión V1 no se ofrece asesoría automática sobre tareas o instalaciones («{q}»). "
            "Por favor, busque artículos concretos por su nombre, SKU o sinónimo (por ejemplo: «Martillo», «Tornillo», «Sellador»)."
        ),
        "out_of_scope_title": "ℹ️ Consulta fuera del alcance de ferretería",
        "out_of_scope_desc": (
            "Solo puedo ayudarte a consultar productos, existencias y ubicaciones del catálogo de ferretería y bricolaje."
        ),
        "sheet_header": "📋 Ficha del Producto",
        "field_sku": "Código SKU",
        "field_name": "Producto",
        "field_category": "Categoría",
        "field_brand": "Marca",
        "field_measure": "Medida",
        "field_unit": "Unidad de venta",
        "field_stock": "Existencias (Stock)",
        "field_status": "Estado de disponibilidad",
        "field_description": "Descripción del catálogo",
        "field_synonyms": "Sinónimos registrados",
        "field_specs": "Especificaciones técnicas (según catálogo)",
        "specs_disclaimer": (
            "Nota: Los datos técnicos proceden exclusivamente de las celdas registradas en inventario.xlsx "
            "y no constituyen asesoría técnica experta ni recomendación automática de compatibilidad."
        ),
        "location_header": "📍 Ubicación Física en Tienda",
        "location_rack": "Rack",
        "location_column": "Columna",
        "location_row": "Fila",
        "location_summary": "Rack {rack} / Columna {col} / Fila {row}",
        "out_of_stock_location_msg": (
            "Artículo actualmente AGOTADO (Stock = 0). No se confirma ubicación física disponible en estantería."
        ),
        "not_specified": "No especificado en catálogo",
        "error_loading_title": "⚠️ No se pudo cargar o validar el inventario",
        "error_loading_hint": (
            "Por favor, verifique que el archivo `inventario.xlsx` exista junto a `app.py`, contenga la hoja "
            "`Inventario` con las 13 columnas obligatorias y cumpla las reglas de datos. Luego pulse «Recargar inventario»."
        ),
        "total_items_caption": "Total de referencias cargadas en `inventario.xlsx`: {n} artículos.",
        "match_badge_exact": "Coincidencia exacta",
        "match_badge_partial": "Coincidencia parcial",
        "match_badge_approx": "Posible coincidencia (confirmar)",
    },
    "en": {
        "lang_label": "🌐 Idioma / Language",
        "app_title": "🛠️ FerreteríaMS AI — Product Locator",
        "app_subtitle": "Quickly check stock availability and physical store location from the official catalog.",
        "v1_badge": "ℹ️ Trial version (V1 without AI): AI advisory is not available in this trial version. Search queries read directly from inventario.xlsx.",
        "reload_btn": "🔄 Reload inventory",
        "reload_help": "Re-reads the local inventario.xlsx file and re-runs data validation.",
        "reload_ok": "Inventory successfully reloaded and validated.",
        "tab_search": "🔍 Catalog Search",
        "tab_ai_disabled": "🤖 AI Advisory (Unavailable in V1)",
        "ai_disabled_title": "AI Advisory: not available in this trial version",
        "ai_disabled_desc": (
            "This first version (V1) validates the user interface, SKU/name/synonym search, "
            "stock checking, and Rack / Column / Row location without connecting to AI models "
            "or simulated chats. To find an item, please use the 'Catalog Search' tab."
        ),
        "search_label": "Search by product name, SKU code, or synonym",
        "search_placeholder": "E.g.: Martillo de peña, FER-001, Tornillo 1/2 pulgada, hammer...",
        "category_label": "Filter by category (optional)",
        "all_categories": "All categories",
        "search_btn": "🔎 Search",
        "clear_btn": "🧹 Clear search / New search",
        "results_header": "Search Results",
        "catalog_header": "Catalog products available for lookup",
        "exact_matches_msg": "Found {n} exact or direct match(es):",
        "partial_matches_msg": "Found {n} match(es) in the catalog:",
        "approx_warning_title": "⚠️ Possible match (approximate spelling suggestion)",
        "approx_warning_desc": (
            "No exact match was found for '{q}'. "
            "Showing possible approximate matches below. Please manually confirm and select the desired item "
            "by verifying its SKU and measure:"
        ),
        "select_product_prompt": "Select an item from the list to view its full product sheet and location:",
        "view_sheet_btn": "View sheet for {sku}",
        "selected_btn": "✓ Viewing sheet for {sku}",
        "status_available": "AVAILABLE",
        "status_out_of_stock": "OUT OF STOCK",
        "status_not_found": "NOT FOUND",
        "not_found_title": "❌ Product not found",
        "not_found_desc": (
            "No item matching '{q}' was found in the demo catalog. "
            "Please check the SKU code, product name, or try simpler search terms."
        ),
        "task_query_title": "ℹ️ Technical advisory not available in this trial version",
        "task_query_desc": (
            "This V1 version does not provide automated advice for tasks or installations ('{q}'). "
            "Please search for specific items by name, SKU, or synonym (for example: 'hammer', 'bolt', 'sealant')."
        ),
        "out_of_scope_title": "ℹ️ Query outside hardware store scope",
        "out_of_scope_desc": (
            "I can only help you look up products, stock availability, and locations in the hardware and DIY catalog."
        ),
        "sheet_header": "📋 Product Sheet",
        "field_sku": "SKU Code",
        "field_name": "Product",
        "field_category": "Category",
        "field_brand": "Brand",
        "field_measure": "Measure / Size",
        "field_unit": "Sales Unit",
        "field_stock": "Stock Quantity",
        "field_status": "Availability Status",
        "field_description": "Catalog Description",
        "field_synonyms": "Registered Synonyms",
        "field_specs": "Technical Specifications (from catalog)",
        "specs_disclaimer": (
            "Note: Technical specifications come exclusively from the cells recorded in inventario.xlsx "
            "and do not constitute expert technical advice or automated compatibility recommendations."
        ),
        "location_header": "📍 Physical Store Location",
        "location_rack": "Rack",
        "location_column": "Column",
        "location_row": "Row",
        "location_summary": "Rack {rack} / Column {col} / Row {row}",
        "out_of_stock_location_msg": (
            "Item currently OUT OF STOCK (Stock = 0). No current physical shelf location is confirmed."
        ),
        "not_specified": "Not specified in catalog",
        "error_loading_title": "⚠️ Could not load or validate inventory",
        "error_loading_hint": (
            "Please verify that `inventario.xlsx` exists next to `app.py`, contains the `Inventario` sheet "
            "with all 13 required columns, and satisfies all data rules. Then click 'Reload inventory'."
        ),
        "total_items_caption": "Total references loaded from `inventario.xlsx`: {n} items.",
        "match_badge_exact": "Exact match",
        "match_badge_partial": "Partial match",
        "match_badge_approx": "Possible match (please confirm)",
    },
}


class ErrorValidacionInventario(ValueError):
    """Excepción controlada para errores de estructura o datos en inventario.xlsx."""


# =====================================================================
# 2. LECTURA Y VALIDACIÓN DEL INVENTARIO
# =====================================================================


def limpiar_valor_celda(valor: Any) -> str:
    """Convierte un valor de celda a texto limpio sin producir nunca 'nan'."""
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() == "nan":
        return ""
    return texto


def validar_dataframe_inventario(df: pd.DataFrame) -> pd.DataFrame:
    """Valida las 13 columnas obligatorias y las reglas de integridad de inventario.xlsx.

    Reglas validadas:
    - Presencia exacta de las 13 columnas requeridas.
    - Al menos un registro válido.
    - SKU obligatorio, no vacío y único.
    - Producto y Categoria obligatorios y no vacíos.
    - Stock obligatorio, entero >= 0 (sin negativos, textos ni decimales).
    - Rack, Columna y Fila obligatorios cuando Stock > 0, conservando ceros iniciales.
    """
    if not isinstance(df, pd.DataFrame):
        raise ErrorValidacionInventario("El contenido leído no es una tabla válida.")

    columnas_presentes = list(df.columns)
    columnas_faltantes = [c for c in COLUMNAS_OBLIGATORIAS if c not in columnas_presentes]
    if columnas_faltantes:
        raise ErrorValidacionInventario(
            f"Faltan columnas obligatorias en la hoja '{HOJA_INVENTARIO}': {', '.join(columnas_faltantes)}."
        )

    if df.empty:
        raise ErrorValidacionInventario(
            f"La hoja '{HOJA_INVENTARIO}' está vacía y no contiene productos."
        )

    df_limpio = df[COLUMNAS_OBLIGATORIAS].copy()

    # Limpiar columnas de texto sin convertir vacíos en 'nan'
    columnas_texto = [c for c in COLUMNAS_OBLIGATORIAS if c != "Stock"]
    for col in columnas_texto:
        df_limpio[col] = df_limpio[col].apply(limpiar_valor_celda)

    skus_vistos: set[str] = set()
    stocks_validados: list[int] = []

    for idx, fila in df_limpio.iterrows():
        num_fila_excel = int(idx) + 2  # +2 considerando encabezado en fila 1
        sku = fila["SKU"]
        producto = fila["Producto"]
        categoria = fila["Categoria"]
        stock_bruto = df.loc[idx, "Stock"]

        if not sku:
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel}: El campo 'SKU' está vacío y es obligatorio."
            )
        sku_norm = sku.upper()
        if sku_norm in skus_vistos:
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel}: El SKU '{sku}' está duplicado. Cada SKU debe ser único."
            )
        skus_vistos.add(sku_norm)

        if not producto:
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel} (SKU {sku}): El campo 'Producto' está vacío y es obligatorio."
            )
        if not categoria:
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel} (SKU {sku}): El campo 'Categoria' está vacío y es obligatorio."
            )

        # Validar Stock entero >= 0
        if stock_bruto is None or pd.isna(stock_bruto) or isinstance(stock_bruto, bool):
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel} (SKU {sku}): El campo 'Stock' está vacío o es inválido."
            )
        try:
            if isinstance(stock_bruto, str):
                texto_stock = stock_bruto.strip()
                if not texto_stock or "." in texto_stock or "," in texto_stock:
                    raise ValueError("Stock con formato decimal o vacío")
                valor_stock = int(texto_stock)
            elif isinstance(stock_bruto, (int, float)):
                if float(stock_bruto) != int(stock_bruto):
                    raise ValueError("Stock no entero")
                valor_stock = int(stock_bruto)
            else:
                valor_stock = int(stock_bruto)
        except (ValueError, TypeError) as exc:
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel} (SKU {sku}): 'Stock' debe ser un número entero mayor o igual a 0 (valor recibido: '{stock_bruto}')."
            ) from exc

        if valor_stock < 0:
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel} (SKU {sku}): 'Stock' no puede ser negativo ({valor_stock})."
            )
        stocks_validados.append(valor_stock)

        # Validar ubicación física cuando Stock > 0
        rack = fila["Rack"]
        columna = fila["Columna"]
        fila_ubic = fila["Fila"]
        if valor_stock > 0 and (not rack or not columna or not fila_ubic):
            raise ErrorValidacionInventario(
                f"Fila {num_fila_excel} (SKU {sku}): Con Stock > 0 ({valor_stock}), los campos 'Rack', 'Columna' y 'Fila' son obligatorios."
            )

    df_limpio["Stock"] = stocks_validados
    return df_limpio.reset_index(drop=True)


def cargar_y_validar_inventario(ruta_archivo: Path = INVENTARIO) -> pd.DataFrame:
    """Lee el archivo Excel desde `ruta_archivo`, verifica la hoja `Inventario` y valida sus datos."""
    ruta_resuelta = Path(ruta_archivo)
    if not ruta_resuelta.exists() or not ruta_resuelta.is_file():
        raise ErrorValidacionInventario(
            f"No se encontró el archivo de inventario en la ruta esperada: {ruta_resuelta.name}."
        )

    try:
        df_bruto = pd.read_excel(
            ruta_resuelta,
            sheet_name=HOJA_INVENTARIO,
            engine="openpyxl",
            dtype={"SKU": str, "Rack": str, "Columna": str, "Fila": str},
        )
    except ValueError as exc:
        raise ErrorValidacionInventario(
            f"No se encontró la hoja obligatoria '{HOJA_INVENTARIO}' en '{ruta_resuelta.name}'."
        ) from exc
    except Exception as exc:
        raise ErrorValidacionInventario(
            f"No fue posible abrir o leer el archivo '{ruta_resuelta.name}'. Verifique que sea un archivo Excel (.xlsx) válido y no esté dañado."
        ) from exc

    return validar_dataframe_inventario(df_bruto)


def obtener_firma_archivo(ruta_archivo: Path = INVENTARIO) -> tuple[int, int]:
    """Obtiene (mtime_ns, st_size) del archivo para invalidar automáticamente la caché cuando cambia en disco."""
    ruta_resuelta = Path(ruta_archivo)
    try:
        if ruta_resuelta.exists() and ruta_resuelta.is_file():
            stat_info = ruta_resuelta.stat()
            return (stat_info.st_mtime_ns, stat_info.st_size)
    except OSError:
        pass
    return (0, 0)


@st.cache_data(show_spinner=False, max_entries=2)
def obtener_inventario_cacheado(
    ruta_str: str,
    mtime_ns: int = 0,
    tamano_bytes: int = 0,
) -> pd.DataFrame:
    """Versión en caché de la lectura de inventario para Streamlit.

    Incluye `mtime_ns` y `tamano_bytes` en la clave de caché para detectar
    automáticamente cualquier actualización del archivo `inventario.xlsx` en disco.
    """
    _ = (mtime_ns, tamano_bytes)
    return cargar_y_validar_inventario(Path(ruta_str))


def recargar_inventario() -> None:
    """Limpia la caché de datos de Streamlit para forzar una nueva lectura de inventario.xlsx."""
    obtener_inventario_cacheado.clear()
    st.cache_data.clear()


# =====================================================================
# 3. NORMALIZACIÓN, CLASIFICACIÓN Y MOTOR DE BÚSQUEDA LOCAL
# =====================================================================


def normalizar_texto(texto: Any) -> str:
    """Normaliza texto para búsquedas: minúsculas, sin tildes/diacríticos y espacios simples."""
    limpio = limpiar_valor_celda(texto).lower()
    if not limpio:
        return ""
    descompuesto = unicodedata.normalize("NFD", limpio)
    sin_tildes = "".join(c for c in descompuesto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sin_tildes).strip()


def extraer_lista_sinonimos(cadena_sinonimos: str) -> list[str]:
    """Divide la columna Sinonimos por ';' y limpia espacios extremos."""
    texto = limpiar_valor_celda(cadena_sinonimos)
    if not texto:
        return []
    return [item.strip() for item in texto.split(";") if item.strip()]


PALABRAS_TAREA: tuple[str, ...] = (
    "necesito",
    "como puedo",
    "como hacer",
    "ayuda para",
    "quiero reparar",
    "quiero instalar",
    "quiero fijar",
    "sujetar",
    "pegar",
    "arreglar",
    "instalar",
    "reparar",
    "recomienda",
    "recomendacion",
    "que uso para",
    "que necesito para",
    "i need to",
    "how to",
    "how do i",
    "help me fix",
    "what should i use",
)

PALABRAS_FUERA_ALCANCE: tuple[str, ...] = (
    "capital de",
    "presidente",
    "clima",
    "receta",
    "futbol",
    "pelicula",
    "cancion",
    "chiste",
    "poema",
    "historia de",
    "quien gano",
    "quien es",
    "cual es la capital",
    "what is the capital",
    "weather",
    "recipe",
)


def clasificar_consulta_sin_resultados(consulta: str) -> str:
    """Clasifica una consulta sin coincidencias en:
    - 'fuera_de_alcance': preguntas generales ajenas a ferretería.
    - 'tarea': descripciones de un problema/tarea que requerirían asesoría.
    - 'no_encontrado': búsqueda normal de un producto inexistente.
    """
    q_norm = normalizar_texto(consulta)
    if not q_norm:
        return "vacia"

    for patron in PALABRAS_FUERA_ALCANCE:
        if patron in q_norm:
            return "fuera_de_alcance"

    for patron in PALABRAS_TAREA:
        if patron in q_norm:
            return "tarea"

    return "no_encontrado"


def buscar_productos(
    df: pd.DataFrame,
    consulta: str,
    categoria: str | None = None,
) -> dict[str, Any]:
    """Busca productos en el DataFrame validado según SKU, Producto y Sinonimos.

    Prioridad:
    1. Coincidencia exacta (`exacta`): SKU exacto, nombre de producto exacto o sinónimo exacto.
    2. Coincidencia parcial (`parcial`): la consulta es subcadena de SKU, Producto o Sinonimos,
       o todos los tokens de la consulta aparecen en Producto/Medida/Sinonimos del mismo SKU.
    3. Coincidencia aproximada (`aproximada`): sugerencia ortográfica con `difflib` cuando
       no hay coincidencias exactas ni parciales. Requiere confirmación del usuario.
    """
    df_base = df.copy()
    if categoria and categoria.strip():
        cat_norm = normalizar_texto(categoria)
        df_base = df_base[
            df_base["Categoria"].apply(normalizar_texto) == cat_norm
        ].reset_index(drop=True)

    q_norm = normalizar_texto(consulta)
    if not q_norm:
        return {
            "tipo_resultado": "catalogo",
            "requiere_confirmacion": False,
            "resultados": df_base.to_dict(orient="records"),
        }

    exactos: list[dict[str, Any]] = []
    parciales: list[dict[str, Any]] = []
    tokens_consulta = [t for t in q_norm.split(" ") if t]

    for _, fila in df_base.iterrows():
        registro = fila.to_dict()
        sku_norm = normalizar_texto(registro["SKU"])
        prod_norm = normalizar_texto(registro["Producto"])
        medida_norm = normalizar_texto(registro["Medida"])
        sinonimos_norm = [normalizar_texto(s) for s in extraer_lista_sinonimos(registro["Sinonimos"])]

        # 1. ¿Es coincidencia exacta en SKU, Producto o algún Sinónimo individual?
        es_exacto_directo = (
            q_norm == sku_norm
            or q_norm == prod_norm
            or q_norm in sinonimos_norm
        )
        if es_exacto_directo:
            item = dict(registro)
            item["_tipo_coincidencia"] = "exacta"
            item["_puntaje"] = 1.0
            exactos.append(item)
            continue

        # 2. ¿Es coincidencia parcial en SKU, Producto o algún Sinónimo?
        texto_buscable = " ".join([sku_norm, prod_norm, medida_norm, " ".join(sinonimos_norm)]).strip()
        es_subcadena = (
            q_norm in sku_norm
            or q_norm in prod_norm
            or any(q_norm in s for s in sinonimos_norm)
        )
        # Coincidencia por todos los tokens si todos aparecen como palabras o subcadenas claras
        todos_tokens_presentes = (
            len(tokens_consulta) > 1
            and all(tok in texto_buscable.split() or tok in prod_norm for tok in tokens_consulta)
        )

        if es_subcadena or todos_tokens_presentes:
            item = dict(registro)
            item["_tipo_coincidencia"] = "parcial"
            item["_puntaje"] = 0.85
            parciales.append(item)

    # Si hay exactos: si también hay parciales (ej. "martillo" es sinónimo exacto en FER-001 y parcial en FER-004),
    # combinarlos manteniendo exactos primero y sin duplicar SKU para no ocultar variantes parecidos.
    if exactos or parciales:
        combinados: list[dict[str, Any]] = []
        skus_incluidos: set[str] = set()
        for item in exactos + parciales:
            if item["SKU"] not in skus_incluidos:
                combinados.append(item)
                skus_incluidos.add(item["SKU"])
        tipo = "exacta" if (exactos and len(combinados) == 1) else ("exacta_multiple" if exactos else "parcial")
        return {
            "tipo_resultado": tipo,
            "requiere_confirmacion": False,
            "resultados": combinados,
        }

    # 3. Aproximación prudente con difflib para errores ortográficos (ej. "martiyo")
    # Nota: Si la consulta incluye números/medidas explícitas que no existen (ej. "Tornillo de 3 pulgadas"),
    # NO sugerir otro tamaño distinto como aproximación para no confundir medidas.
    numeros_consulta = set(re.findall(r"\d+(?:/\d+)?", q_norm))

    aproximados: list[dict[str, Any]] = []
    umbral_aproximacion = 0.72

    for _, fila in df_base.iterrows():
        registro = fila.to_dict()
        sku_norm = normalizar_texto(registro["SKU"])
        prod_norm = normalizar_texto(registro["Producto"])
        medida_norm = normalizar_texto(registro["Medida"])
        sinonimos_norm = [normalizar_texto(s) for s in extraer_lista_sinonimos(registro["Sinonimos"])]

        # Si el usuario pidió una medida numérica concreta y el producto tiene otra medida numérica distinta, no aproximar
        if numeros_consulta:
            numeros_producto = set(re.findall(r"\d+(?:/\d+)?", f"{sku_norm} {prod_norm} {medida_norm}"))
            if not numeros_consulta.issubset(numeros_producto):
                continue

        candidatos_texto: list[str] = [sku_norm, prod_norm] + sinonimos_norm
        palabras_producto = [p for p in prod_norm.split(" ") if len(p) >= 4]
        for s in sinonimos_norm:
            palabras_producto.extend([p for p in s.split(" ") if len(p) >= 4])

        mejor_ratio = 0.0
        for cand in candidatos_texto:
            if not cand:
                continue
            ratio = difflib.SequenceMatcher(None, q_norm, cand).ratio()
            if ratio > mejor_ratio:
                mejor_ratio = ratio

        # Comparar también palabra por palabra para consultas de una sola palabra con typo (ej. "martiyo" -> "martillo")
        if len(tokens_consulta) == 1 and len(q_norm) >= 4:
            for palabra in palabras_producto:
                ratio_palabra = difflib.SequenceMatcher(None, q_norm, palabra).ratio()
                if ratio_palabra > mejor_ratio:
                    mejor_ratio = ratio_palabra

        if mejor_ratio >= umbral_aproximacion:
            item = dict(registro)
            item["_tipo_coincidencia"] = "aproximada"
            item["_puntaje"] = round(mejor_ratio, 3)
            aproximados.append(item)

    aproximados.sort(key=lambda x: x["_puntaje"], reverse=True)

    if aproximados:
        return {
            "tipo_resultado": "aproximada",
            "requiere_confirmacion": True,
            "resultados": aproximados,
        }

    # 4. Sin coincidencias: distinguir tarea, fuera de alcance o producto no encontrado
    clasificacion = clasificar_consulta_sin_resultados(consulta)
    return {
        "tipo_resultado": clasificacion,
        "requiere_confirmacion": False,
        "resultados": [],
    }


# =====================================================================
# 4. LÓGICA DE DISPONIBILIDAD Y UBICACIÓN
# =====================================================================


def obtener_estado_producto(producto: dict[str, Any] | None, idioma: str = "es") -> dict[str, Any]:
    """Determina el estado obligatorio ('disponible', 'agotado' o 'no encontrado') y la ubicación.

    Reglas:
    - 'disponible' si existe el SKU y Stock > 0 -> incluye Rack, Columna y Fila.
    - 'agotado' si existe el SKU y Stock == 0 -> NO afirma ubicación física actual disponible.
    - 'no encontrado' si `producto` es None -> sin datos ni ubicación inventada.
    """
    textos = I18N.get(idioma, I18N["es"])
    if not producto or not limpiar_valor_celda(producto.get("SKU")):
        return {
            "estado": "no encontrado",
            "etiqueta_estado": textos["status_not_found"],
            "disponible": False,
            "stock": 0,
            "mostrar_ubicacion": False,
            "ubicacion_formateada": "",
            "rack": "",
            "columna": "",
            "fila": "",
        }

    stock = int(producto.get("Stock", 0))
    if stock > 0:
        rack = limpiar_valor_celda(producto.get("Rack"))
        columna = limpiar_valor_celda(producto.get("Columna"))
        fila = limpiar_valor_celda(producto.get("Fila"))
        ubicacion = textos["location_summary"].format(rack=rack, col=columna, row=fila)
        return {
            "estado": "disponible",
            "etiqueta_estado": textos["status_available"],
            "disponible": True,
            "stock": stock,
            "mostrar_ubicacion": True,
            "ubicacion_formateada": ubicacion,
            "rack": rack,
            "columna": columna,
            "fila": fila,
        }

    return {
        "estado": "agotado",
        "etiqueta_estado": textos["status_out_of_stock"],
        "disponible": False,
        "stock": 0,
        "mostrar_ubicacion": False,
        "ubicacion_formateada": "",
        "rack": "",
        "columna": "",
        "fila": "",
    }


# =====================================================================
# 5. COMPONENTES DE INTERFAZ Y VISUALIZACIÓN EN STREAMLIT
# =====================================================================


def inicializar_estado_sesion() -> None:
    """Inicializa las claves necesarias en `st.session_state` sin crear historial de chat."""
    if "idioma" not in st.session_state:
        st.session_state.idioma = "es"
    if "consulta_actual" not in st.session_state:
        st.session_state.consulta_actual = ""
    if "input_busqueda" not in st.session_state:
        st.session_state.input_busqueda = ""
    if "categoria_actual" not in st.session_state:
        st.session_state.categoria_actual = ""
    if "sku_seleccionado" not in st.session_state:
        st.session_state.sku_seleccionado = None
    if "aviso_recarga" not in st.session_state:
        st.session_state.aviso_recarga = False


def limpiar_busqueda_callback() -> None:
    """Restablece la consulta, el filtro de categoría y la selección de SKU conservando el idioma."""
    st.session_state.consulta_actual = ""
    st.session_state.input_busqueda = ""
    st.session_state.categoria_actual = ""
    st.session_state.sku_seleccionado = None


def seleccionar_sku_callback(sku: str) -> None:
    """Guarda en `st.session_state` el SKU seleccionado por el usuario."""
    st.session_state.sku_seleccionado = sku


def inyectar_estilos_responsivos() -> None:
    """Aplica estilos CSS ligeros para destacar ubicación y estados en móvil y escritorio."""
    st.markdown(
        """
        <style>
        .v1-banner {
            padding: 0.75rem 1rem;
            border-radius: 0.5rem;
            background-color: #f0f4f8;
            border-left: 5px solid #0284c7;
            color: #0f172a;
            margin-bottom: 1rem;
            font-size: 0.95rem;
        }
        .status-badge-avail {
            display: inline-block;
            padding: 0.25rem 0.7rem;
            border-radius: 999px;
            background-color: #dcfce7;
            color: #166534;
            font-weight: 700;
            border: 1px solid #86efac;
        }
        .status-badge-out {
            display: inline-block;
            padding: 0.25rem 0.7rem;
            border-radius: 999px;
            background-color: #fee2e2;
            color: #991b1b;
            font-weight: 700;
            border: 1px solid #fca5a5;
        }
        .location-box {
            padding: 1rem 1.25rem;
            border-radius: 0.6rem;
            background-color: #eff6ff;
            border: 2px solid #2563eb;
            color: #1e3a8a;
            margin: 0.75rem 0;
        }
        .location-box h4 {
            margin: 0 0 0.4rem 0;
            color: #1d4ed8;
        }
        .location-coords {
            font-size: 1.35rem;
            font-weight: 800;
            letter-spacing: 0.02em;
        }
        @media (max-width: 640px) {
            .location-coords {
                font-size: 1.1rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def renderizar_ficha_producto(producto: dict[str, Any], idioma: str) -> None:
    """Muestra la ficha detallada de un producto validado del Excel sin inventar datos."""
    t = I18N[idioma]
    info_estado = obtener_estado_producto(producto, idioma=idioma)

    with st.container(border=True):
        st.subheader(f"{t['sheet_header']}: {producto['Producto']}")

        # Cabecera de identificación y estado en texto claro + estilo visual
        col_id, col_estado = st.columns([2, 1])
        with col_id:
            st.markdown(
                f"**{t['field_sku']}:** `{producto['SKU']}` &nbsp;|&nbsp; "
                f"**{t['field_category']}:** {producto['Categoria']}"
            )
            medida_txt = producto["Medida"] if producto["Medida"] else t["not_specified"]
            st.markdown(f"**{t['field_measure']}:** `{medida_txt}`")
        with col_estado:
            clase_badge = "status-badge-avail" if info_estado["disponible"] else "status-badge-out"
            icono_estado = "✅" if info_estado["disponible"] else "⛔"
            st.markdown(
                f"<div><strong>{t['field_status']}:</strong><br>"
                f"<span class='{clase_badge}'>{icono_estado} {info_estado['etiqueta_estado']}</span></div>",
                unsafe_allow_html=True,
            )

        st.divider()

        # Bloque destacado de Ubicación Física (solo si Stock > 0) o aviso de Agotado
        if info_estado["mostrar_ubicacion"]:
            st.markdown(
                f"""
                <div class="location-box">
                    <h4>{t['location_header']}</h4>
                    <div class="location-coords">{info_estado['ubicacion_formateada']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            c_rack, c_col, c_fila, c_stock = st.columns(4)
            c_rack.metric(t["location_rack"], info_estado["rack"])
            c_col.metric(t["location_column"], info_estado["columna"])
            c_fila.metric(t["location_row"], info_estado["fila"])
            unidad_txt = f" {producto['Unidad']}" if producto.get("Unidad") else ""
            c_stock.metric(t["field_stock"], f"{info_estado['stock']}{unidad_txt}")
        else:
            st.warning(f"⛔ **{info_estado['etiqueta_estado']}** — {t['out_of_stock_location_msg']}")
            c_stock, c_med = st.columns(2)
            c_stock.metric(t["field_stock"], "0")
            c_med.metric(
                t["field_measure"],
                producto["Medida"] if producto["Medida"] else t["not_specified"],
            )

        st.divider()

        # Campos descriptivos tal como constan en el Excel
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            marca_txt = producto["Marca"] if producto["Marca"] else t["not_specified"]
            unidad_val = producto["Unidad"] if producto["Unidad"] else t["not_specified"]
            st.markdown(f"**{t['field_brand']}:** {marca_txt}")
            st.markdown(f"**{t['field_unit']}:** {unidad_val}")
        with col_d2:
            sinonimos_txt = producto["Sinonimos"] if producto["Sinonimos"] else t["not_specified"]
            st.markdown(f"**{t['field_synonyms']}:** {sinonimos_txt}")

        desc_txt = producto["Descripcion"] if producto["Descripcion"] else t["not_specified"]
        st.markdown(f"**{t['field_description']}:**\n\n{desc_txt}")

        espec_txt = (
            producto["Especificaciones_tecnicas"]
            if producto["Especificaciones_tecnicas"]
            else t["not_specified"]
        )
        st.markdown(f"**{t['field_specs']}:**\n\n{espec_txt}")
        st.caption(t["specs_disclaimer"])


def renderizar_lista_resultados(
    resultados: list[dict[str, Any]],
    requiere_confirmacion: bool,
    idioma: str,
) -> None:
    """Muestra la lista de productos coincidentes diferenciando SKU, Medida y Estado."""
    t = I18N[idioma]

    # Autoseleccionar únicamente cuando hay 1 resultado exacto/parcial que NO sea aproximación ortográfica
    if len(resultados) == 1 and not requiere_confirmacion:
        st.session_state.sku_seleccionado = resultados[0]["SKU"]

    if len(resultados) > 1 or requiere_confirmacion:
        st.info(t["select_product_prompt"])

    for item in resultados:
        sku = item["SKU"]
        info_estado = obtener_estado_producto(item, idioma=idioma)
        medida = item["Medida"] if item["Medida"] else t["not_specified"]
        tipo_coincidencia = item.get("_tipo_coincidencia", "")

        etiqueta_tipo = ""
        if tipo_coincidencia == "exacta":
            etiqueta_tipo = f" · 🎯 {t['match_badge_exact']}"
        elif tipo_coincidencia == "parcial":
            etiqueta_tipo = f" · 🔹 {t['match_badge_partial']}"
        elif tipo_coincidencia == "aproximada":
            etiqueta_tipo = f" · ⚠️ {t['match_badge_approx']}"

        icono = "✅" if info_estado["disponible"] else "⛔"
        ubic_resumen = (
            f" | 📍 {info_estado['ubicacion_formateada']}"
            if info_estado["mostrar_ubicacion"]
            else ""
        )

        with st.container(border=True):
            c_info, c_accion = st.columns([3, 1])
            with c_info:
                st.markdown(
                    f"**`{sku}` — {item['Producto']}** ({t['field_measure']}: `{medida}`){etiqueta_tipo}"
                )
                st.caption(
                    f"{t['field_category']}: {item['Categoria']} | "
                    f"{icono} **{info_estado['etiqueta_estado']}** ({t['field_stock']}: {info_estado['stock']})"
                    f"{ubic_resumen}"
                )
            with c_accion:
                es_seleccionado = st.session_state.sku_seleccionado == sku
                texto_boton = (
                    t["selected_btn"].format(sku=sku)
                    if es_seleccionado
                    else t["view_sheet_btn"].format(sku=sku)
                )
                st.button(
                    texto_boton,
                    key=f"btn_select_{sku}",
                    type="primary" if es_seleccionado else "secondary",
                    use_container_width=True,
                    on_click=seleccionar_sku_callback,
                    args=(sku,),
                )


def main() -> None:
    """Punto de entrada principal de la aplicación Streamlit FerreteríaMS AI V1."""
    st.set_page_config(
        page_title="FerreteríaMS AI · V1",
        page_icon="🛠️",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    inicializar_estado_sesion()
    inyectar_estilos_responsivos()

    # Barra superior: Selector de idioma y botón de recarga de inventario
    col_top_title, col_top_lang, col_top_reload = st.columns([3, 1.2, 1.2])

    with col_top_lang:
        opciones_idioma = {"es": "Español (ES)", "en": "English (EN)"}
        idioma_elegido = st.selectbox(
            I18N[st.session_state.idioma]["lang_label"],
            options=list(opciones_idioma.keys()),
            format_func=lambda k: opciones_idioma[k],
            index=0 if st.session_state.idioma == "es" else 1,
            key="selector_idioma_widget",
        )
        if idioma_elegido != st.session_state.idioma:
            st.session_state.idioma = idioma_elegido
            st.rerun()

    t = I18N[st.session_state.idioma]

    with col_top_title:
        st.title(t["app_title"])
        st.caption(t["app_subtitle"])

    with col_top_reload:
        st.write("")  # Espaciado vertical para alinear con el selector
        if st.button(t["reload_btn"], help=t["reload_help"], use_container_width=True):
            recargar_inventario()
            st.session_state.aviso_recarga = True
            st.rerun()

    if st.session_state.aviso_recarga:
        st.toast(t["reload_ok"], icon="✅")
        st.session_state.aviso_recarga = False

    # Banner obligatorio de versión de prueba sin IA
    st.markdown(f"<div class='v1-banner'>{t['v1_badge']}</div>", unsafe_allow_html=True)

    # Cargar y validar el inventario antes de habilitar cualquier búsqueda
    try:
        mtime_ns, tamano_bytes = obtener_firma_archivo(INVENTARIO)
        df_inventario = obtener_inventario_cacheado(
            str(INVENTARIO),
            mtime_ns=mtime_ns,
            tamano_bytes=tamano_bytes,
        )
    except ErrorValidacionInventario as err_val:
        st.error(f"**{t['error_loading_title']}**\n\n{err_val}")
        st.info(t["error_loading_hint"])
        return
    except Exception:
        st.error(f"**{t['error_loading_title']}**")
        st.info(t["error_loading_hint"])
        return

    # Pestañas: Buscador activo y pestaña de IA explícitamente deshabilitada
    tab_buscador, tab_ia = st.tabs([t["tab_search"], t["tab_ai_disabled"]])

    with tab_ia:
        st.warning(f"🔒 **{t['ai_disabled_title']}**")
        st.write(t["ai_disabled_desc"])

    with tab_buscador:
        # Controles de búsqueda y filtro por categoría
        categorias_disponibles = sorted(df_inventario["Categoria"].unique().tolist())
        opciones_cat = [""] + categorias_disponibles

        col_busq, col_cat = st.columns([2.5, 1.2])
        with col_busq:
            entrada_usuario = st.text_input(
                t["search_label"],
                key="input_busqueda",
                placeholder=t["search_placeholder"],
            )
        with col_cat:
            indice_cat = (
                opciones_cat.index(st.session_state.categoria_actual)
                if st.session_state.categoria_actual in opciones_cat
                else 0
            )
            cat_seleccionada = st.selectbox(
                t["category_label"],
                options=opciones_cat,
                index=indice_cat,
                format_func=lambda c: t["all_categories"] if c == "" else c,
            )

        col_btn_buscar, col_btn_limpiar, _ = st.columns([1, 1.3, 2.7])
        with col_btn_buscar:
            ejecutar_busqueda = st.button(
                t["search_btn"], type="primary", use_container_width=True
            )
        with col_btn_limpiar:
            st.button(
                t["clear_btn"],
                on_click=limpiar_busqueda_callback,
                use_container_width=True,
            )

        # Actualizar estado de búsqueda cuando cambia el input o el filtro
        if ejecutar_busqueda or entrada_usuario != st.session_state.consulta_actual:
            st.session_state.consulta_actual = entrada_usuario
            st.session_state.sku_seleccionado = None
        if cat_seleccionada != st.session_state.categoria_actual:
            st.session_state.categoria_actual = cat_seleccionada
            st.session_state.sku_seleccionado = None

        st.divider()

        # Ejecutar motor de búsqueda local
        respuesta_busqueda = buscar_productos(
            df_inventario,
            consulta=st.session_state.consulta_actual,
            categoria=st.session_state.categoria_actual,
        )
        tipo_res = respuesta_busqueda["tipo_resultado"]
        lista_res = respuesta_busqueda["resultados"]
        requiere_conf = respuesta_busqueda["requiere_confirmacion"]

        # Mostrar mensajes según el tipo de resultado
        if tipo_res == "fuera_de_alcance":
            st.warning(
                f"**{t['out_of_scope_title']}**\n\n{t['out_of_scope_desc']}"
            )
        elif tipo_res == "tarea":
            st.info(
                f"**{t['task_query_title']}**\n\n"
                f"{t['task_query_desc'].format(q=st.session_state.consulta_actual)}"
            )
        elif tipo_res == "no_encontrado":
            st.error(
                f"**{t['not_found_title']} ({t['status_not_found']})**\n\n"
                f"{t['not_found_desc'].format(q=st.session_state.consulta_actual)}"
            )
        elif tipo_res == "aproximada":
            st.warning(
                f"**{t['approx_warning_title']}**\n\n"
                f"{t['approx_warning_desc'].format(q=st.session_state.consulta_actual)}"
            )
            renderizar_lista_resultados(lista_res, requiere_confirmacion=True, idioma=st.session_state.idioma)
        elif tipo_res in ("exacta", "exacta_multiple", "parcial"):
            st.subheader(t["results_header"])
            msg_conteo = (
                t["exact_matches_msg"].format(n=len(lista_res))
                if tipo_res == "exacta"
                else t["partial_matches_msg"].format(n=len(lista_res))
            )
            st.caption(msg_conteo)
            renderizar_lista_resultados(lista_res, requiere_confirmacion=False, idioma=st.session_state.idioma)
        else:
            st.subheader(t["catalog_header"])
            st.caption(t["total_items_caption"].format(n=len(lista_res)))
            renderizar_lista_resultados(lista_res, requiere_confirmacion=False, idioma=st.session_state.idioma)

        # Si hay un SKU seleccionado que pertenece a los resultados actuales, mostrar su ficha completa
        if st.session_state.sku_seleccionado:
            producto_sel = next(
                (item for item in lista_res if item["SKU"] == st.session_state.sku_seleccionado),
                None,
            )
            if producto_sel is not None:
                st.divider()
                renderizar_ficha_producto(producto_sel, idioma=st.session_state.idioma)


if __name__ == "__main__":
    main()