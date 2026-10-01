"""Testes unitários para o módulo normalizer.py."""

from __future__ import annotations

import unittest

from src.core.normalizer import (
    extrair_cnaes,
    formatar_cnae,
    formatar_cep,
    formatar_cnpj,
    montar_endereco_completo,
    normalizar_bairro,
    normalizar_cep,
    normalizar_cnpj,
)


class TestNormalizer(unittest.TestCase):
    """Bateria de testes para higienização e formatação."""

    def test_normalizar_cnpj(self) -> None:
        self.assertEqual(normalizar_cnpj("61.017.032/0001-94"), "61017032000194")
        self.assertEqual(normalizar_cnpj("61017032000194"), "61017032000194")
        self.assertEqual(normalizar_cnpj("7244271000131"), "07244271000131")  # zfill 14
        self.assertEqual(normalizar_cnpj("12345678901"), "00012345678901")  # zfill 14 para CNPJs truncados
        self.assertEqual(normalizar_cnpj(""), "")
        self.assertEqual(normalizar_cnpj(None), "")
        self.assertEqual(normalizar_cnpj("abc"), "")

    def test_formatar_cnpj(self) -> None:
        self.assertEqual(formatar_cnpj("61017032000194"), "61.017.032/0001-94")
        self.assertEqual(formatar_cnpj("07244271000131"), "07.244.271/0001-31")
        self.assertEqual(formatar_cnpj(""), "")
        self.assertEqual(formatar_cnpj(None), "")

    def test_normalizar_cep(self) -> None:
        self.assertEqual(normalizar_cep("15050-090"), "15050090")
        self.assertEqual(normalizar_cep("15050090"), "15050090")
        self.assertEqual(normalizar_cep("5050090"), "05050090")  # zfill 8
        self.assertEqual(normalizar_cep(""), "")
        self.assertEqual(normalizar_cep(None), "")

    def test_formatar_cep(self) -> None:
        self.assertEqual(formatar_cep("15050090"), "15050-090")
        self.assertEqual(formatar_cep(""), "")
        self.assertEqual(formatar_cep(None), "")

    def test_normalizar_bairro(self) -> None:
        # Acentos e minúsculas
        self.assertEqual(normalizar_bairro("Vila Anchieta"), "VILA ANCHIETA")
        self.assertEqual(normalizar_bairro("SÃO FRANCISCO"), "SAO FRANCISCO")
        self.assertEqual(normalizar_bairro("ELDORADO (2ª PARTE)"), "ELDORADO 2A PARTE")

        # Abreviações consagradas
        self.assertEqual(normalizar_bairro("Jd. Tarraf"), "JARDIM TARRAF")
        self.assertEqual(normalizar_bairro("JD MARACANÃ"), "JARDIM MARACANA")
        self.assertEqual(normalizar_bairro("Pq. da Liberdade"), "PARQUE DA LIBERDADE")
        self.assertEqual(normalizar_bairro("Vl. Imperial"), "VILA IMPERIAL")
        self.assertEqual(normalizar_bairro("Res. Palestra"), "RESIDENCIAL PALESTRA")
        self.assertEqual(normalizar_bairro("Conj. Habitacional"), "CONJUNTO HABITACIONAL")
        self.assertEqual(normalizar_bairro("Cj. Habitacional"), "CONJUNTO HABITACIONAL")
        self.assertEqual(normalizar_bairro("Chac. Jockey Club"), "CHACARA JOCKEY CLUB")
        self.assertEqual(normalizar_bairro("Est. Jockey Club"), "ESTANCIA JOCKEY CLUB")
        self.assertEqual(normalizar_bairro("Cond. Recanto Real"), "CONDOMINIO RECANTO REAL")
        self.assertEqual(normalizar_bairro("Dist. Ind. Waldemar"), "DISTRITO INDUSTRIAL WALDEMAR")

        # Não corrupção de algarismos romanos (V, IV, etc.)
        self.assertEqual(normalizar_bairro("DAMHA V"), "DAMHA V")
        self.assertEqual(normalizar_bairro("PARQUE DA LIBERDADE V"), "PARQUE DA LIBERDADE V")
        self.assertEqual(normalizar_bairro("AUFERVILLE V"), "AUFERVILLE V")
        self.assertEqual(normalizar_bairro("JARDIM TANGARA J"), "JARDIM TANGARA J")

        # Casos nulos / vazios
        self.assertEqual(normalizar_bairro(""), "")
        self.assertEqual(normalizar_bairro(None), "")

    def test_extrair_cnaes(self) -> None:
        cnaes_raw = "461410004, 461410005, 461500001, 749010401"
        resultado = extrair_cnaes(cnaes_raw)
        self.assertEqual(resultado, ["4614100", "4615000", "7490104"])

        # Ponto e vírgula
        cnaes_sc = "472370000; 561120100"
        self.assertEqual(extrair_cnaes(cnaes_sc), ["4723700", "5611201"])

        # Vazio ou nulo
        self.assertEqual(extrair_cnaes(""), [])
        self.assertEqual(extrair_cnaes(None), [])

    def test_formatar_cnae(self) -> None:
        self.assertEqual(formatar_cnae("4687701"), "4687-7/01")
        self.assertEqual(formatar_cnae("0111301"), "0111-3/01")
        self.assertEqual(formatar_cnae(""), "")
        self.assertEqual(formatar_cnae(None), "")

    def test_montar_endereco_completo(self) -> None:
        # Endereço com dois-pontos no complemento vindo do VRE
        end1 = montar_endereco_completo("RUA", "AIMORES", "287", ":", "VILA ANCHIETA")
        self.assertEqual(end1, "RUA AIMORES, 287 - VILA ANCHIETA")

        # Endereço com complemento real
        end2 = montar_endereco_completo("AV", "MÁRIO ANDREAZZA", "181", "BRCAO 02", "JARDIM SÃO MARCO")
        self.assertEqual(end2, "AV MÁRIO ANDREAZZA, 181 BRCAO 02 - JARDIM SÃO MARCO")

        # Sem número
        end3 = montar_endereco_completo("RUA", "CENTRAL", "", "", "CENTRO")
        self.assertEqual(end3, "RUA CENTRAL - CENTRO")


if __name__ == "__main__":
    unittest.main()
