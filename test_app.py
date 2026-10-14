"""Pruebas automatizadas para FerreteríaMS AI · Versión 1 (V1 sin IA).

Verifica:
1. Lectura real de `inventario.xlsx` con ruta portable (`Path(__file__).resolve().parent`).
2. Validación de las 13 columnas obligatorias y reglas de integridad de datos.
3. Casos de prueba obligatorios definidos en `AGENTS.md` y `especificacion.md`.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd

import app


class TestFerreteriaMSV1(unittest.TestCase):
    """Suite de pruebas para lectura, validación, búsqueda, disponibilidad e idioma."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.df = app.cargar_y_validar_inventario(app.INVENTARIO)

    def test_01_lectura_y_13_columnas_inventario_real(self) -> None:
        """Verifica que inventario.xlsx abre, contiene las 13 columnas y conserva ceros iniciales."""
        self.assertEqual(list(self.df.columns), app.COLUMNAS_OBLIGATORIAS)
        self.assertGreaterEqual(len(self.df), 8)

        # Verificar que ninguna celda contenga el texto 'nan'
        for col in app.COLUMNAS_OBLIGATORIAS:
            if col != "Stock":
                valores = self.df[col].tolist()
                for val in valores:
                    self.assertNotEqual(str(val).lower(), "nan")

        # Verificar que Fila conserva ceros iniciales ('03', '02', '01')
        fer001 = self.df[self.df["SKU"] == "FER-001"].iloc[0]
        self.assertEqual(fer001["Fila"], "03")
        self.assertEqual(fer001["Rack"], "R02")
        self.assertEqual(fer001["Columna"], "A")

    def test_02_caso_martillo_de_pena(self) -> None:
        """Caso 1: 'Martillo de peña' -> FER-001, disponible, Rack R02 / Columna A / Fila 03."""
        res = app.buscar_productos(self.df, "Martillo de peña")
        self.assertIn(res["tipo_resultado"], ("exacta", "exacta_multiple", "parcial"))
        self.assertFalse(res["requiere_confirmacion"])
        self.assertEqual(res["resultados"][0]["SKU"], "FER-001")

        estado = app.obtener_estado_producto(res["resultados"][0], idioma="es")
        self.assertEqual(estado["estado"], "disponible")
        self.assertTrue(estado["mostrar_ubicacion"])
        self.assertEqual(estado["ubicacion_formateada"], "Rack R02 / Columna A / Fila 03")

    def test_03_caso_tornillo_media_pulgada(self) -> None:
        """Caso 2: 'Tornillo 1/2 pulgada' -> FER-002, Rack R05 / Columna B / Fila 02; sin confundir medidas."""
        res = app.buscar_productos(self.df, "Tornillo 1/2 pulgada")
        self.assertEqual(res["tipo_resultado"], "exacta")
        self.assertEqual(len(res["resultados"]), 1)
        prod = res["resultados"][0]
        self.assertEqual(prod["SKU"], "FER-002")
        self.assertEqual(prod["Medida"], "1/2 pulgada")

        estado = app.obtener_estado_producto(prod, idioma="es")
        self.assertEqual(estado["estado"], "disponible")
        self.assertEqual(estado["ubicacion_formateada"], "Rack R05 / Columna B / Fila 02")

    def test_04_caso_error_ortografico_martiyo(self) -> None:
        """Caso 3: 'martiyo' -> posible coincidencia para selección manual, sin autoselección automática."""
        res = app.buscar_productos(self.df, "martiyo")
        self.assertEqual(res["tipo_resultado"], "aproximada")
        self.assertTrue(res["requiere_confirmacion"])
        skus_sugeridos = [r["SKU"] for r in res["resultados"]]
        self.assertIn("FER-001", skus_sugeridos)

    def test_05_caso_producto_agotado_fer005(self) -> None:
        """Caso 4: 'FER-005' -> agotado (Stock = 0), sin prometer ubicación física actual."""
        res = app.buscar_productos(self.df, "FER-005")
        self.assertEqual(res["tipo_resultado"], "exacta")
        prod = res["resultados"][0]
        self.assertEqual(prod["SKU"], "FER-005")

        estado = app.obtener_estado_producto(prod, idioma="es")
        self.assertEqual(estado["estado"], "agotado")
        self.assertEqual(estado["stock"], 0)
        self.assertFalse(estado["mostrar_ubicacion"])
        self.assertEqual(estado["ubicacion_formateada"], "")

    def test_06_caso_no_encontrado_y_medida_inexistente(self) -> None:
        """Caso 5: SKU inexistente o medida no registrada ('Tornillo de 3 pulgadas') -> no encontrado."""
        res_sku = app.buscar_productos(self.df, "FER-999")
        self.assertEqual(res_sku["tipo_resultado"], "no_encontrado")
        self.assertEqual(len(res_sku["resultados"]), 0)

        res_medida = app.buscar_productos(self.df, "Tornillo de 3 pulgadas")
        self.assertEqual(res_medida["tipo_resultado"], "no_encontrado")
        self.assertEqual(len(res_medida["resultados"]), 0)

        estado_vacio = app.obtener_estado_producto(None, idioma="es")
        self.assertEqual(estado_vacio["estado"], "no encontrado")
        self.assertFalse(estado_vacio["mostrar_ubicacion"])

    def test_07_caso_busqueda_sinonimos_ingles_y_etiquetas(self) -> None:
        """Caso 6: Búsqueda en inglés mediante sinónimo ('hammer', 'safety glasses') y etiquetas en inglés."""
        res_hammer = app.buscar_productos(self.df, "hammer")
        skus = [r["SKU"] for r in res_hammer["resultados"]]
        self.assertIn("FER-001", skus)
        self.assertIn("FER-004", skus)

        res_glasses = app.buscar_productos(self.df, "safety glasses")
        self.assertEqual(res_glasses["resultados"][0]["SKU"], "FER-007")

        estado_en = app.obtener_estado_producto(res_glasses["resultados"][0], idioma="en")
        self.assertEqual(estado_en["etiqueta_estado"], "AVAILABLE")
        self.assertEqual(estado_en["ubicacion_formateada"], "Rack R01 / Column D / Row 01")

    def test_08_caso_varios_articulos_parecidos(self) -> None:
        """Caso 7: Varios artículos parecidos ('martillo') -> devuelve FER-001 y FER-004 diferenciados."""
        res = app.buscar_productos(self.df, "martillo")
        self.assertGreaterEqual(len(res["resultados"]), 2)
        skus = [r["SKU"] for r in res["resultados"]]
        self.assertIn("FER-001", skus)
        self.assertIn("FER-004", skus)

    def test_09_consultas_tarea_y_fuera_de_alcance(self) -> None:
        """Verifica clasificación de consultas de tarea ('Necesito sujetar una tubería') y fuera de alcance."""
        res_tarea = app.buscar_productos(self.df, "Necesito sujetar una tubería")
        self.assertEqual(res_tarea["tipo_resultado"], "tarea")
        self.assertEqual(len(res_tarea["resultados"]), 0)

        res_fuera = app.buscar_productos(self.df, "¿Cuál es la capital de Francia?")
        self.assertEqual(res_fuera["tipo_resultado"], "fuera_de_alcance")
        self.assertEqual(len(res_fuera["resultados"]), 0)

    def test_10_validaciones_errores_archivo_y_datos(self) -> None:
        """Caso 9: Archivo ausente/roto, columna faltante, SKU duplicado, stock inválido o ubicación incompleta."""
        # 1. Archivo inexistente
        with self.assertRaises(app.ErrorValidacionInventario):
            app.cargar_y_validar_inventario(Path("archivo_que_no_existe_12345.xlsx"))

        # 2. Archivo corrupto
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_corrupta = Path(tmpdir) / "corrupto.xlsx"
            ruta_corrupta.write_text("esto no es un zip excel valido", encoding="utf-8")
            with self.assertRaises(app.ErrorValidacionInventario):
                app.cargar_y_validar_inventario(ruta_corrupta)

        # 3. Columna faltante
        df_sin_col = self.df.drop(columns=["Medida"])
        with self.assertRaises(app.ErrorValidacionInventario):
            app.validar_dataframe_inventario(df_sin_col)

        # 4. SKU duplicado
        df_sku_dup = self.df.copy()
        df_sku_dup.loc[1, "SKU"] = df_sku_dup.loc[0, "SKU"]
        with self.assertRaises(app.ErrorValidacionInventario):
            app.validar_dataframe_inventario(df_sku_dup)

        # 5. Stock negativo
        df_stock_neg = self.df.copy()
        df_stock_neg.loc[0, "Stock"] = -5
        with self.assertRaises(app.ErrorValidacionInventario):
            app.validar_dataframe_inventario(df_stock_neg)

        # 6. Stock no entero
        df_stock_txt = self.df.copy()
        df_stock_txt["Stock"] = df_stock_txt["Stock"].astype(object)
        df_stock_txt.loc[0, "Stock"] = "doce"
        with self.assertRaises(app.ErrorValidacionInventario):
            app.validar_dataframe_inventario(df_stock_txt)

        # 7. Ubicación incompleta con Stock > 0
        df_sin_ubic = self.df.copy()
        df_sin_ubic.loc[0, "Rack"] = ""
        with self.assertRaises(app.ErrorValidacionInventario):
            app.validar_dataframe_inventario(df_sin_ubic)

    def test_11_interfaz_streamlit_apptest(self) -> None:
        """Caso 10, 11 e Issue #1: Verifica el arranque de la interfaz Streamlit, ausencia de ejemplos rápidos, búsqueda, limpieza y cambio de idioma."""
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file("app.py").run()
        self.assertEqual(len(at.exception), 0)
        self.assertIn("Localizador de Productos", at.title[0].value)

        # Issue #1: Verificar que no existe el texto "Ejemplos rápidos de prueba" ni los botones de ejemplo
        textos_captions = [c.value for c in at.caption]
        self.assertFalse(any("Ejemplos rápidos de prueba" in txt for txt in textos_captions))
        self.assertFalse(any("Quick test examples" in txt for txt in textos_captions))
        etiquetas_botones = [b.label for b in at.button]
        for ej in ("`Martillo de peña`", "`Tornillo 1/2 pulgada`", "`FER-005`", "`martiyo`", "`hammer`"):
            self.assertNotIn(ej, etiquetas_botones)

        # 1. Búsqueda exacta 'Martillo de peña' -> muestra ficha de FER-001
        at.text_input(key="input_busqueda").set_value("Martillo de peña").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["sku_seleccionado"], "FER-001")

        # 2. Búsqueda aproximada 'martiyo' -> no autoselecciona SKU y muestra advertencia de posible coincidencia
        at.text_input(key="input_busqueda").set_value("martiyo").run()
        self.assertEqual(len(at.exception), 0)
        self.assertIsNone(at.session_state["sku_seleccionado"])
        self.assertTrue(any("Posible coincidencia" in w.value for w in at.warning))

        # 3. Cambiar idioma a inglés -> conserva estado sin errores y traduce interfaz
        at.selectbox(key="selector_idioma_widget").set_value("en").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["idioma"], "en")
        self.assertIn("Product Locator", at.title[0].value)


if __name__ == "__main__":
    unittest.main()
