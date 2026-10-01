"""Testes de integridade da Interface Gráfica Administrativa (Tkinter).

Verifica que a classe PainelAdminVRE instancia todos os widgets
(árvores, botões, formulários, dropdowns e ordenação) sem gerar AttributeError.
"""

from __future__ import annotations

import unittest
from unittest import mock


class TestAdminGuiIntegridade(unittest.TestCase):
    """Testa a integridade estrutural da interface Tkinter."""

    def test_instanciacao_painel_admin_widgets(self):
        """Verifica se PainelAdminVRE cria todos os atributos de widgets esperados."""
        fake_tk = mock.MagicMock()
        fake_ttk = mock.MagicMock()
        fake_mb = mock.MagicMock()

        modules_mock = {
            "tkinter": fake_tk,
            "tkinter.ttk": fake_ttk,
            "tkinter.messagebox": fake_mb,
        }

        with mock.patch.dict("sys.modules", modules_mock):
            import importlib
            import src.admin.main_admin as ma

            importlib.reload(ma)

            root = mock.MagicMock()
            painel = ma.PainelAdminVRE(root)

            # Verifica widgets essenciais
            self.assertTrue(hasattr(painel, "tree_pend"), "tree_pend deve existir no PainelAdminVRE")
            self.assertTrue(hasattr(painel, "tree_cep"), "tree_cep deve existir no PainelAdminVRE")
            self.assertTrue(hasattr(painel, "tree_hist"), "tree_hist deve existir no PainelAdminVRE")
            self.assertTrue(hasattr(painel, "tree_lotes"), "tree_lotes deve existir no PainelAdminVRE")
            self.assertTrue(hasattr(painel, "ent_form_cep"), "ent_form_cep deve existir")
            self.assertTrue(hasattr(painel, "cb_form_area"), "cb_form_area deve existir")
            self.assertTrue(hasattr(painel, "ent_form_bairro"), "ent_form_bairro deve existir")

            # Verifica novos dropdowns solicitados pelo usuário
            self.assertTrue(hasattr(painel, "cb_filtro_fiscal"), "cb_filtro_fiscal deve existir")
            self.assertTrue(hasattr(painel, "cb_filtro_area"), "cb_filtro_area deve existir")
            self.assertTrue(hasattr(painel, "cb_filtro_bairro"), "cb_filtro_bairro deve existir")

            # Verifica métodos de carregamento sem congelamento e ordenação
            self.assertTrue(hasattr(painel, "_executar_com_loading"), "_executar_com_loading deve existir")
            self.assertTrue(hasattr(painel, "ordenar_treeview"), "ordenar_treeview deve existir")
            self.assertTrue(hasattr(painel, "carregar_ultimos_fiscais"), "carregar_ultimos_fiscais deve existir")

            # Verifica método de obtenção de áreas
            areas = painel._obter_areas_disponiveis()
            self.assertIsInstance(areas, list)
            self.assertIn("CENTRAL", areas)
            self.assertIn("HB", areas)


if __name__ == "__main__":
    unittest.main()
