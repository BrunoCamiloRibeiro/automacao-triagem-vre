"""Gerador de Arquivos de Teste Abrangentes para a Triagem VRE (SMAURB).

Gera múltiplos arquivos CSVs cobrindo todos os cenários operacionais e de borda:
1. Empresas novas sem histórico
2. Empresas com histórico de fiscal prévio
3. CNAEs Noturnos (incluindo herança hierárquica: restaurantes, bares, danceterias)
4. CNAEs com Restrição (frigoríficos, reciclagem, sucata)
5. Noturno e Restrição combinados simultaneamente
6. Resolução de Área via fallback de CEP e registro de CEPs pendentes
7. Encodings e caracteres acentuados da língua portuguesa (UTF-8 e Latin-1)
8. Teste de carga e volume com 500 protocolos
"""

from __future__ import annotations

import csv
import random
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "tests" / "test_data"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CABECALHOS_VRE = [
    "ProtocoloSil",
    "ProtocoloLicenca",
    "Data",
    "GrauRiscoGeral",
    "SituacaoPostura",
    "InscricaoMunicipal",
    "Cnpj",
    "NomeEmpresarial",
    "Telefone",
    "E-mail",
    "TipoEndereco",
    "Logradouro",
    "Numero",
    "Complemento",
    "Bairro",
    "Cep",
    "FinalidadeEndereco",
    "CnaesLicenciadas",
    "CnaesNaoLicenciadas",
    "Iptu",
    "",
]


def _gravar_csv(caminho: Path, linhas: list[list[str]], encoding: str = "latin-1") -> None:
    with open(caminho, mode="w", encoding=encoding, newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(CABECALHOS_VRE)
        for linha in linhas:
            # Garante que tenha 21 colunas terminando em vazio (padrão do sistema VRE)
            l = list(linha)
            while len(l) < 20:
                l.append("")
            if len(l) == 20:
                l.append("")
            writer.writerow(l)
    print(f"[✓] Gerado: {caminho.name} ({len(linhas)} linhas, encoding={encoding})")


def gerar_cenario_01_novas():
    """Empresas novas sem fiscal, sem restrição, horário diurno."""
    linhas = [
        [
            "5241901", "A0010001000100010001000100010001", "10/06/2026", "BAIXO", "DEFERIDO", "901001",
            "11.222.333/0001-44", "NOVA CONSULTORIA E SERVICOS ADMINISTRATIVOS LTDA",
            "17-32001111", "contato@novaconsultoria.com.br", "RUA", "BERNARDINO DE CAMPOS", "3000", "SALA 1",
            "CENTRO", "15015300", "EXERCICIO DA ATIVIDADE", "702040001, 821999901", "", "0102030405"
        ],
        [
            "5241902", "A0010001000100010001000100010002", "10/06/2026", "MEDIO", "DEFERIDO", "901002",
            "22.333.444/0001-55", "RIO PRETO VESTUARIO E MODA EIRELI",
            "17-32002222", "financeiro@modariopreto.com.br", "AV", "BADY BASSITT", "4500", "",
            "BOA VISTA", "15025000", "EXERCICIO DA ATIVIDADE", "478140001", "", "0203040506"
        ],
    ]
    _gravar_csv(OUTPUT_DIR / "01_teste_empresas_novas.csv", linhas)


def gerar_cenario_02_historico_fiscais():
    """Empresas com histórico de atendimento prévio no SQLite."""
    linhas = [
        # 33.424.347/0001-10 -> Fiscal_Alfa
        [
            "5241910", "A0020002000200020002000200020001", "11/06/2026", "MEDIO", "DEFERIDO", "4225170",
            "33.424.347/0001-10", "EMPRESA EXEMPLO ALFA LTDA",
            "11-94195652", "contato@empresa-alfa.com.br", "RUA", "AIMORES", "287", "",
            "VILA ANCHIETA", "15050090", "EXERCICIO DA ATIVIDADE", "461410004", "", "0307968000"
        ],
        # 42.454.244/0001-94 -> Fiscal_Beta
        [
            "5241911", "A0020002000200020002000200020002", "11/06/2026", "ALTO", "DEFERIDO", "4035240",
            "42.454.244/0001-94", "EMPRESA EXEMPLO BETA S/A",
            "17-33334444", "contato@empresa-beta.com.br", "AV", "NOSSA SENHORA DA PAZ", "1200", "",
            "JARDIM AMERICA", "15055000", "EXERCICIO DA ATIVIDADE", "251100001", "", "0405060708"
        ],
    ]
    _gravar_csv(OUTPUT_DIR / "02_teste_historico_fiscais.csv", linhas)


def gerar_cenario_03_cnaes_noturnos():
    """Empresas com atividades noturnas e herança hierárquica de CNAEs."""
    linhas = [
        # Restaurante e Similares: Subclasse 5611-2/01 (Herdado de Classe 56.11-2 / Grupo 56.1)
        [
            "5241920", "A0030003000300030003000300030001", "12/06/2026", "ALTO", "DEFERIDO", "902001",
            "33.444.555/0001-66", "CHURRASCARIA E RESTAURANTE ESTRELA DO SUL LTDA",
            "17-32115555", "contato@estreladosul.com.br", "AV", "ALBERTO ANDALO", "3500", "",
            "CENTRO", "15015000", "EXERCICIO DA ATIVIDADE", "561120101", "", "0506070809"
        ],
        # Bar: Subclasse 5611-2/02
        [
            "5241921", "A0030003000300030003000300030002", "12/06/2026", "MEDIO", "DEFERIDO", "902002",
            "44.555.666/0001-77", "BOTECO DOS AMIGOS CHOPERIA E PETISCOS EIRELI",
            "17-32226666", "boteco@amigoschopp.com.br", "RUA", "REDENTORA", "2800", "",
            "VILA REDENTORA", "15015750", "EXERCICIO DA ATIVIDADE", "561120201", "", "0607080910"
        ],
        # Danceteria / Salão de Dança: 9329-8/01
        [
            "5241922", "A0030003000300030003000300030003", "12/06/2026", "ALTO", "DEFERIDO", "902003",
            "55.666.777/0001-88", "BOATE E DANCETERIA ILUSION CLUB SHOWS LTDA",
            "17-32337777", "eventos@ilusionclub.com.br", "AV", "JOSÉ MUNIA", "5000", "",
            "JARDIM REDENTORA", "15090500", "EXERCICIO DA ATIVIDADE", "932980101", "", "0708091011"
        ],
    ]
    _gravar_csv(OUTPUT_DIR / "03_teste_cnaes_noturnos_heranca.csv", linhas)


def gerar_cenario_04_cnaes_restricao():
    """Empresas com atividades de fiscalização complexa / restrição."""
    linhas = [
        # Abate de bovinos / Frigorífico: 1011-2/01 (Restrição = 1)
        [
            "5241930", "A0040004000400040004000400040001", "13/06/2026", "ALTO", "DEFERIDO", "903001",
            "66.777.888/0001-99", "FRIGORIFICO CENTRAL DE CARNES DE RIO PRETO LTDA",
            "17-32448888", "carnes@frigocentral.com.br", "AV", "PHILADELPHO MANOEL GOUVEIA NETTO", "1800", "",
            "PARQUE INDUSTRIAL", "15030760", "EXERCICIO DA ATIVIDADE", "101120101", "", "0809101112"
        ],
        # Fabricação de produtos de carne: 1013-9/01 (Restrição = 1)
        [
            "5241931", "A0040004000400040004000400040002", "13/06/2026", "ALTO", "DEFERIDO", "903002",
            "77.888.999/0001-10", "EMBUTIDOS E CARNES NOROESTE S/A",
            "17-32559999", "contato@embutidos.com.br", "ESTRADA", "VICINAL TALHADO", "KM 4", "",
            "TALHADO", "15100000", "EXERCICIO DA ATIVIDADE", "101390101", "", "0910111213"
        ],
    ]
    _gravar_csv(OUTPUT_DIR / "04_teste_cnaes_restricao.csv", linhas)


def gerar_cenario_05_combinados():
    """Empresas que possuem simultaneamente restrição e atividade noturna."""
    linhas = [
        # Frigorífico / Carnes (Restrição 1011-2/01) + Restaurante (Noturno 5611-2/01)
        [
            "5241940", "A0050005000500050005000500050001", "14/06/2026", "ALTO", "DEFERIDO", "904001",
            "88.999.000/0001-21", "FRIGORIFICO BOI NA BRASA CHURRASCARIA NOTURNA LTDA",
            "17-32660000", "diretoria@boinabrasa.com.br", "RODOVIA", "WASHINGTON LUIS", "KM 435", "",
            "DISTRITO INDUSTRIAL", "15080000", "EXERCICIO DA ATIVIDADE", "101120101, 561120101", "", "1011121314"
        ]
    ]
    _gravar_csv(OUTPUT_DIR / "05_teste_noturno_e_restricao_combinados.csv", linhas)


def gerar_cenario_06_fallback_cep():
    """Resolução de Área via Fallback de CEP e registro de pendentes."""
    linhas = [
        # Linha 1: Bairro inexistente/em branco, mas CEP mapeado (15042015 -> CIDADE DA CRIANÇA)
        [
            "5241950", "A0060006000600060006000600060001", "15/06/2026", "MEDIO", "DEFERIDO", "905001",
            "99.000.111/0001-32", "COMERCIO DE BEBIDAS DO BAIRRO DESCONHECIDO EIRELI",
            "17-32771111", "bebidas@desconhecido.com.br", "RUA", "PROJETADA A", "100", "",
            "", "15042015", "EXERCICIO DA ATIVIDADE", "472370001", "", "1112131415"
        ],
        # Linha 2: Bairro desconhecido E CEP desconhecido (deve ir para cep_pendentes)
        [
            "5241951", "A0060006000600060006000600060002", "15/06/2026", "BAIXO", "DEFERIDO", "905002",
            "10.111.222/0001-43", "EMPREENDIMENTO NOVO LOTEAMENTO PENDENTE LTDA",
            "17-32882222", "contato@lotependente.com.br", "RUA", "NOVA EXPEDICAO", "50", "",
            "BAIRRO FICTICIO INEXISTENTE ZZZ", "15099888", "EXERCICIO DA ATIVIDADE", "681020101", "", "1213141516"
        ],
    ]
    _gravar_csv(OUTPUT_DIR / "06_teste_fallback_cep_e_pendencias.csv", linhas)


def gerar_cenario_07_encodings():
    """Arquivo em UTF-8 com acentuações e caracteres especiais."""
    linhas = [
        [
            "5241960", "A0070007000700070007000700070001", "16/06/2026", "MEDIO", "DEFERIDO", "906001",
            "21.222.333/0001-54", "AÇOUGUE E MERCEARIA SÃO JOÃO BATISTA DE RIO PRETO",
            "17-32993333", "saojoao@acougue.com.br", "AV", "SÃO FRANCISCO", "750", "2ª PARTE",
            "ELDORADO (2ª PARTE)", "15043000", "EXERCICIO DA ATIVIDADE", "472290101", "", "1314151617"
        ],
        [
            "5241961", "A0070007000700070007000700070002", "16/06/2026", "BAIXO", "DEFERIDO", "906002",
            "32.333.444/0001-65", "CONFECÇÕES & COSTURA CONCEIÇÃO APARECIDA ME",
            "17-32004444", "costura@conceicao.com.br", "RUA", "JOÃO PAULO II", "120", "",
            "CONJUNTO HABITACIONAL SÃO JOSÉ", "15040000", "EXERCICIO DA ATIVIDADE", "141260101", "", "1415161718"
        ],
    ]
    _gravar_csv(OUTPUT_DIR / "07_teste_encodings_e_acentuacao.csv", linhas, encoding="utf-8")


def gerar_cenario_08_volume_500():
    """Simulação de lote volumoso com 500 registros para benchmark de estabilidade."""
    bairros = ["CENTRO", "BOA VISTA", "VILA ANCHIETA", "VILA REDENTORA", "JARDIM AMERICA", "ELDORADO"]
    logradouros = ["RUA VOLUNTARIOS DE SAO PAULO", "AV BADY BASSITT", "AV ALBERTO ANDALO", "RUA BERNARDINO DE CAMPOS"]
    cnaes_amostra = [
        "471130201",  # Diurno comum
        "561120101",  # Noturno
        "468770101",  # Restrição
        "561120201",  # Noturno
        "702040001",  # Diurno
    ]

    linhas = []
    for i in range(1, 501):
        prot_sil = str(5300000 + i)
        prot_lic = f"B{i:04d}000000000000000000000000{i:04d}"
        cnpj_num = f"{random.randint(10, 99)}.{random.randint(100, 999)}.{random.randint(100, 999)}/0001-{random.randint(10, 99)}"
        razao = f"EMPRESA DE TESTE DE CARGA {i:04d} LTDA"
        bairro = random.choice(bairros)
        logr = random.choice(logradouros)
        num = str(random.randint(10, 5000))
        cnae = random.choice(cnaes_amostra)

        linhas.append(
            [
                prot_sil, prot_lic, "20/06/2026", "MEDIO", "DEFERIDO", str(800000 + i),
                cnpj_num, razao, "17-30000000", f"teste{i}@cargavre.com.br", "RUA", logr, num, "",
                bairro, "15015000", "EXERCICIO DA ATIVIDADE", cnae, "", f"00{i:08d}"
            ]
        )
    _gravar_csv(OUTPUT_DIR / "08_teste_carga_volume_500_linhas.csv", linhas)


def gerar_todos():
    print("[*] Gerando bateria de arquivos de teste para o VRE...")
    gerar_cenario_01_novas()
    gerar_cenario_02_historico_fiscais()
    gerar_cenario_03_cnaes_noturnos()
    gerar_cenario_04_cnaes_restricao()
    gerar_cenario_05_combinados()
    gerar_cenario_06_fallback_cep()
    gerar_cenario_07_encodings()
    gerar_cenario_08_volume_500()
    print("[✓] Todos os 8 cenários foram gerados com sucesso em:", OUTPUT_DIR)


if __name__ == "__main__":
    gerar_todos()
