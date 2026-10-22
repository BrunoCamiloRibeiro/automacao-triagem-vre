"""Testes de casos de borda, entradas vazias e resiliência a erros."""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from src.core.cep_service import (
    consultar_area_por_cep,
    ensure_cep_tables,
    listar_ceps_pendentes,
    registrar_area_cep,
    registrar_cep_pendente,
    resolver_area,
)
from src.core.database import connect_db, initialize_schema_if_needed
from src.core.normalizer import (
    extrair_cnaes,
    formatar_cep,
    formatar_cnae,
    formatar_cnpj,
    montar_endereco_completo,
    normalizar_bairro,
    normalizar_cep,
    normalizar_cnpj,
)
from src.runner.main_seco import (
    extrair_mes_competencia,
    limpar_csvs_processados,
    processar_fluxo_seco,
)


class TestCasosDeBorda(unittest.TestCase):
    """Testa casos limite, valores nulos, formatos corrompidos e boundary conditions."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.db_t = self.root / "triagem.db"
        self.db_c = self.root / "cep.db"
        self.conn_t = connect_db(self.db_t)
        self.conn_c = connect_db(self.db_c)
        initialize_schema_if_needed(self.conn_t, is_cep=False)
        ensure_cep_tables(self.conn_c)

    def tearDown(self) -> None:
        self.conn_t.close()
        self.conn_c.close()
        self.tmp_dir.cleanup()

    def test_normalizadores_entradas_estranhas(self) -> None:
        # CNPJ com caracteres especiais e espaços
        self.assertEqual(normalizar_cnpj("  12.345.678/0001-90 \n"), "12345678000190")
        self.assertEqual(normalizar_cnpj("---"), "")
        self.assertEqual(normalizar_cnpj(12345678000190), "12345678000190")

        # CEP com letras misturadas
        self.assertEqual(normalizar_cep("CEP: 15.010-000"), "15010000")
        self.assertEqual(normalizar_cep(15010000), "15010000")
        self.assertEqual(normalizar_cep(""), "")

        # CNAEs com formatos bizarros ou incompletos
        self.assertEqual(extrair_cnaes("123, 4567, 1234567890"), ["1234567"])
        self.assertEqual(extrair_cnaes(""), [])
        self.assertEqual(extrair_cnaes("   "), [])
        self.assertEqual(extrair_cnaes("abc, def"), [])

        # Bairros com pontuação mista e espaços repetidos
        self.assertEqual(normalizar_bairro("   Jd.   São   José   -  2   "), "JARDIM SAO JOSE 2")
        self.assertEqual(normalizar_bairro(""), "")
        self.assertEqual(normalizar_bairro("   "), "")

        # Endereço sem número e sem complemento
        self.assertEqual(montar_endereco_completo("RUA", "TESTE", None, None, None), "RUA TESTE")
        self.assertEqual(montar_endereco_completo("", "", "", ":", ""), "")

    def test_extrair_mes_competencia(self) -> None:
        # Formatos válidos
        self.assertEqual(extrair_mes_competencia("01/05/2026"), "2026-05")
        self.assertEqual(extrair_mes_competencia("2026-05-01"), "2026-05")
        self.assertEqual(extrair_mes_competencia("15/12/2024"), "2024-12")

        # Formatos inválidos devem cair no fallback do mês corrente sem lançar exceção
        comp_vazio = extrair_mes_competencia("")
        self.assertRegex(comp_vazio, r"^\d{4}-\d{2}$")

        comp_invalido = extrair_mes_competencia("data_invalida")
        self.assertRegex(comp_invalido, r"^\d{4}-\d{2}$")

    def test_cep_service_boundary_values(self) -> None:
        # CEP inválido deve disparar ValueError ao cadastrar
        with self.assertRaises(ValueError):
            registrar_area_cep("123", "CENTRAL", conn=self.conn_c)

        # Área vazia deve disparar ValueError
        with self.assertRaises(ValueError):
            registrar_area_cep("15000000", "", conn=self.conn_c)

        # Consulta com CEP malformatado retorna None
        self.assertIsNone(consultar_area_por_cep("abc", conn=self.conn_c))
        self.assertIsNone(consultar_area_por_cep(None, conn=self.conn_c))

        # Registro de pendência com CEP inválido é ignorado silenciosamente
        registrar_cep_pendente("123", conn=self.conn_c)
        pendentes = listar_ceps_pendentes(conn=self.conn_c)
        self.assertEqual(len(pendentes), 0)

    def test_csv_vazio_ou_corrompido(self) -> None:
        csv_vazio = self.root / "vazio.csv"
        csv_vazio.write_text("", encoding="latin-1")

        saida = self.root / "saida.xlsx"
        template = self.root / "template.xlsx"

        # Execução sobre CSV vazio deve concluir retornando 0 registros sem crashar
        total = processar_fluxo_seco(
            caminho_csv=csv_vazio,
            saida_xlsx=saida,
            silencioso=True,
            triagem_db_path=self.db_t,
            cep_db_path=self.db_c,
            template_path=template,
        )
        self.assertEqual(total, 0)
        self.assertTrue(saida.exists())

    def test_utf8_e_latin1_encoding_sem_mojibake(self) -> None:
        """Verifica que CSVs gravados em UTF-8 são decodificados sem mojibake."""
        from src.runner.main_seco import ler_linhas_csv
        csv_utf8 = self.root / "utf8_test.csv"
        conteudo = (
            "ProtocoloLicenca;Cnpj;NomeEmpresarial;Telefone;Email;TipoLogradouro;Logradouro;Numero;"
            "Complemento;Bairro;Cep;DescricaoAtividade;CnaesSecundarias;CnaePrincipal;Viabilidade\n"
            "LIC_UTF8;11222333000144;AÇOUGUE E PADARIA SÃO JOSÉ;1799991111;contato@sj.com;RUA;JOÃO PAULO;10;"
            ";JARDIM SÃO FRANCISCO;15080000;COMÉRCIO DE CARNES;472370000;472110200;123456\n"
        )
        csv_utf8.write_text(conteudo, encoding="utf-8")

        dados = ler_linhas_csv(csv_utf8)
        self.assertEqual(len(dados), 1)
        self.assertEqual(dados[0]["NomeEmpresarial"], "AÇOUGUE E PADARIA SÃO JOSÉ")
        self.assertEqual(dados[0]["Bairro"], "JARDIM SÃO FRANCISCO")

    def test_backfill_area_ao_mapear_cep(self) -> None:
        """Verifica se cadastrar área para CEP atualiza backlog ativo pendente automaticamente."""
        import json
        cur = self.conn_t.cursor()
        cur.execute(
            """
            INSERT INTO vre_backlog (protocolo_licenca, cnpj_limpo, mes_competencia, data_importacao, status, area, dados_extras_json)
            VALUES ('P_PEND', '12345678000100', '2026-07', '2026-07-01', 'PENDENTE', '', ?);
            """,
            (json.dumps({"Cep": "15099888", "Bairro": "Bairro Novo"}),),
        )
        self.conn_t.commit()

        # O protocolo está com area vazia
        cur.execute("SELECT area FROM vre_backlog WHERE protocolo_licenca = 'P_PEND';")
        self.assertEqual(cur.fetchone()["area"], "")

        # Cadastra o CEP no cep_areas.db com triagem_db informado para o backfill
        ok = registrar_area_cep(
            "15099888",
            "REPRESA",
            "Bairro Novo",
            conn=self.conn_c,
            triagem_conn=self.conn_t,
        )
        self.assertTrue(ok)

        # Verifica se o backlog ativo foi backfilled
        cur.execute("SELECT area FROM vre_backlog WHERE protocolo_licenca = 'P_PEND';")
        self.assertEqual(cur.fetchone()["area"], "REPRESA")

    def test_continuidade_fiscal_backlog_ativo(self) -> None:
        """Verifica se o runner busca fiscal no backlog ativo mesmo sem estar em historico_fiscais."""
        cur = self.conn_t.cursor()
        cur.execute(
            """
            INSERT INTO vre_backlog (protocolo_licenca, cnpj_limpo, mes_competencia, data_importacao, status, fiscal_atribuido)
            VALUES ('P_ANTIGO', '99888777000166', '2026-01', '2026-01-15', 'DISTRIBUIDO', 'Fiscal_Designado');
            """
        )
        self.conn_t.commit()

        # Agora processa novo CSV do mesmo CNPJ para o mês 2026-02
        csv_novo = self.root / "novo.csv"
        conteudo = (
            "ProtocoloLicenca;Cnpj;NomeEmpresarial;Telefone;Email;TipoLogradouro;Logradouro;Numero;"
            "Complemento;Bairro;Cep;DescricaoAtividade;CnaesSecundarias;CnaePrincipal;Viabilidade\n"
            "P_NOVO;99.888.777/0001-66;EMPRESA TESTE;17999;test@test.com;RUA;CENTRAL;1;"
            ";CENTRO;15010000;ATIVIDADE;;561120100;123456\n"
        )
        csv_novo.write_text(conteudo, encoding="latin-1")

        saida = self.root / "saida2.xlsx"
        template = self.root / "template2.xlsx"

        processar_fluxo_seco(
            caminho_csv=csv_novo,
            saida_xlsx=saida,
            silencioso=True,
            triagem_db_path=self.db_t,
            cep_db_path=self.db_c,
            template_path=template,
        )

        cur.execute("SELECT fiscal_atribuido FROM vre_backlog WHERE protocolo_licenca = 'P_NOVO';")
        row = cur.fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row["fiscal_atribuido"], "Fiscal_Designado")

    def test_remocao_csv_apos_processamento_e_limpeza(self) -> None:
        """Verifica se arquivos CSV em pasta de entrada são apagados após a execução e se CSVs legados são limpos."""
        import src.core.config as cfg
        from unittest.mock import patch

        pasta_fake_entrada = self.root / "entrada_fake"
        pasta_fake_entrada.mkdir()
        csv_prod = pasta_fake_entrada / "remessa_producao.csv"
        conteudo = (
            "ProtocoloSil;ProtocoloLicenca;Cnpj;NomeEmpresarial;Telefone;E-mail;TipoEndereco;Logradouro;Numero;"
            "Complemento;Bairro;Cep;FinalidadeEndereco;CnaesNaoLicenciadas;CnaesLicenciadas;Iptu\n"
            "SIL_DEL;LIC_DEL;88.777.666/0001-55;EMPRESA DELETE;17999;del@del.com;RUA;CENTRAL;10;"
            ";CENTRO;15010000;ATIVIDADE;;561120100;123456\n"
        )
        csv_prod.write_text(conteudo, encoding="utf-8")
        self.assertTrue(csv_prod.exists())

        # Cria arquivo legado na pasta de processados
        pasta_fake_proc = self.root / "proc_fake"
        pasta_fake_proc.mkdir()
        csv_legado = pasta_fake_proc / "antigo_2026.csv"
        csv_legado.write_text("legado", encoding="utf-8")

        with patch.object(cfg, "ENTRADA_CSV_DIR", pasta_fake_entrada), \
             patch("src.runner.main_seco.ENTRADA_CSV_DIR", pasta_fake_entrada), \
             patch.object(cfg, "PROCESSADOS_CSV_DIR", pasta_fake_proc), \
             patch("src.runner.main_seco.PROCESSADOS_CSV_DIR", pasta_fake_proc):

            # Executa o processamento
            processar_fluxo_seco(
                caminho_csv=csv_prod,
                saida_xlsx=self.root / "saida_del.xlsx",
                silencioso=True,
                triagem_db_path=self.db_t,
                cep_db_path=self.db_c,
                template_path=self.root / "template_del.xlsx",
            )

            # O arquivo CSV processado NÃO deve mais existir (apagado após a execução)
            self.assertFalse(csv_prod.exists(), "O arquivo CSV processado deve ser apagado após execução")

            # O arquivo CSV legado em processados também deve ter sido limpo
            self.assertFalse(csv_legado.exists(), "Arquivos legados em processados devem ser limpos")


if __name__ == "__main__":
    unittest.main()
