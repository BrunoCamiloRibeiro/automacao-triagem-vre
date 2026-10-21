"""Executável Administrativo GUI (config_vre.exe).

Interface gráfica corporativa em Tkinter nativo para gestão operacional da SMAURB:
1. Aba 1 - Busca de Histórico:
   - Pesquisa ágil de atendimentos por CNPJ, Razão Social ou Fiscal titular.
   - Listagem em tempo real dos atendimentos mais recentes.
   - Importação pontual ou em lote de planilhas históricas via thread em background com modal de progresso.
2. Aba 2 - Gestão de CEP x Área:
   - Mapeamento, edição e exclusão de correlações no banco satélite `cep_areas.db`.
   - Filtros combinados por Área Fiscal e Bairro de referência.
   - Painel de CEPs pendentes ordenados pelo volume de ocorrências.
3. Aba 3 - Arquivamento de Competências:
   - Painel de controle de meses de competência (visibilidade: Ativo vs Arquivado).
   - KPIs gerenciais do backlog (total, pendentes, distribuídos, finalizados, visíveis).
   - Botão atômico de aplicação com backup preventivo datado e regeneração da planilha ativa.
4. Modo CLI / Manutenção:
   - Suporte a comandos de terminal (`--cli`, `--buscar`, `--arquivar`, `--mapear-cep`, etc.)
     para ambientes headless ou scripts automatizados.
"""

from __future__ import annotations

import argparse
import sys
import threading
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.admin.archive_service import (
    alternar_status_competencia,
    aplicar_reorganizacao_planilha,
    arquivar_competencia,
    buscar_historico_fiscais,
    definir_status_competencia,
    listar_competencias_status,
    listar_fiscais_cadastrados,
    obter_estatisticas_gerais,
    obter_ultimos_atendimentos,
    reativar_competencia,
)
from src.core.cep_service import (
    listar_areas_mapeadas,
    listar_bairros_mapeados,
    listar_ceps_mapeados,
    listar_ceps_pendentes,
    registrar_area_cep,
    remover_cep_mapeado,
)
from src.core.historico_importer import (
    escanear_e_importar_planilhas_antigas,
    importar_planilha_historico,
)
from src.core.normalizer import formatar_cep, formatar_cnpj, normalizar_cep, normalizar_cnpj

# Carregamento seguro do Tkinter para tolerar ambientes headless (ex: servidores Linux de CI/CD)
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    TKINTER_DISPONIVEL = True
except ImportError:
    TKINTER_DISPONIVEL = False


class PainelAdminVRE:
    """Janela principal do Executável Administrativo da SMAURB / Inspetoria Fiscal."""

    def __init__(self, root: tk.Tk) -> None:
        """Inicializa a interface gráfica, estilos, abas e carrega os dados iniciais.

        Args:
            root: Instância principal da janela Tkinter (`tk.Tk`).
        """
        self.root = root
        self.root.title("SMAURB - Painel Administrativo VRE & Fiscais")
        self.root.geometry("1060x700")
        self.root.minsize(900, 580)

        # Referências de imagens para prevenir garbage collection precoce pelo Tkinter
        self.img_logo: Any = None
        self.img_header: Any = None

        # Configuração de Estilo ttk profissional
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._carregar_logo()
        self._criar_widgets()
        self.atualizar_tudo()

    def _carregar_logo(self) -> None:
        """Carrega o brasão de São José do Rio Preto para o ícone e cabeçalho da janela."""
        candidatos = [
            PROJECT_ROOT / "templates" / "logo_rio_preto.png",
            PROJECT_ROOT / "fbuild" / "Configurações" / "templates" / "logo_rio_preto.png",
            PROJECT_ROOT / "Configurações" / "templates" / "logo_rio_preto.png",
            Path("templates/logo_rio_preto.png").resolve(),
        ]
        for p in candidatos:
            if p.exists() and p.is_file():
                try:
                    self.img_logo = tk.PhotoImage(file=str(p))
                    # Define o brasão como ícone oficial da janela e da barra de tarefas
                    self.root.iconphoto(True, self.img_logo)
                    # Versão redimensionada proporcional para o banner de cabeçalho (~45px)
                    self.img_header = self.img_logo.subsample(2, 2)
                    break
                except Exception:
                    pass

    def _criar_widgets(self) -> None:
        """Constrói a hierarquia de widgets: cabeçalho azul, abas do notebook e status bar."""
        # Cabeçalho Institucional Azul SMAURB (#2C3E50)
        header_frame = tk.Frame(self.root, bg="#2C3E50", height=54)
        header_frame.pack(fill="x", side="top")

        # Exibe o brasão municipal no cabeçalho se disponível
        if self.img_header:
            lbl_logo = tk.Label(header_frame, image=self.img_header, bg="#2C3E50")
            lbl_logo.pack(side="left", padx=(15, 8), pady=4)

        lbl_titulo = tk.Label(
            header_frame,
            text="PREFEITURA DE SÃO JOSÉ DO RIO PRETO — SMAURB / INSPETORIA FISCAL",
            bg="#2C3E50",
            fg="#FFFFFF",
            font=("Arial", 11, "bold"),
            pady=8,
        )
        lbl_titulo.pack(side="left", padx=5)

        # Notebook com as 3 Abas Principais
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # 1. Aba Busca de Histórico
        self.tab_historico = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_historico, text="  🔍 Histórico de Fiscais  ")
        self._construir_aba_historico()

        # 2. Aba CEP x Área
        self.tab_cep = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_cep, text="  🗺️ Gestão de CEP x Área  ")
        self._construir_aba_cep()

        # 3. Aba Arquivamento
        self.tab_arquivamento = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_arquivamento, text="  📦 Arquivamento de Competências  ")
        self._construir_aba_arquivamento()

        # Barra de status inferior
        self.lbl_status = tk.Label(
            self.root,
            text="Sistema pronto para operação.",
            bd=1,
            relief="sunken",
            anchor="w",
            font=("Arial", 9),
            padx=10,
            pady=3,
        )
        self.lbl_status.pack(fill="x", side="bottom")

    def _executar_com_loading(
        self,
        target_fn,
        titulo: str,
        mensagem: str,
        on_sucesso=None,
        on_erro=None,
    ) -> None:
        """Executa uma função pesada em thread de background com modal de progresso indeterminada.

        Garante que a interface gráfica permaneça fluida e responsiva durante operações de I/O
        intensivo (leitura de planilhas gigantes, sincronização do Excel ou geração de relatórios).
        A sincronização com a thread principal do Tkinter é realizada via `root.after(0, ...)`.

        Args:
            target_fn: Função a ser executada na thread de trabalho.
            titulo: Título da janela modal de carregamento.
            mensagem: Texto explicativo exibido para o usuário.
            on_sucesso: Callback opcional chamado com o resultado de `target_fn`.
            on_erro: Callback opcional chamado caso `target_fn` levante exceção.
        """
        top = tk.Toplevel(self.root)
        top.title(titulo)
        top.geometry("460x170")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        # Centraliza o diálogo sobre a janela pai
        try:
            self.root.update_idletasks()
            x = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - 230
            y = self.root.winfo_rooty() + (self.root.winfo_height() // 2) - 85
            top.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        frame = ttk.Frame(top, padding=20)
        frame.pack(fill="both", expand=True)

        lbl_msg = ttk.Label(
            frame,
            text=mensagem,
            font=("Arial", 10, "bold"),
            wraplength=420,
            justify="center",
        )
        lbl_msg.pack(pady=(5, 12))

        pb = ttk.Progressbar(frame, mode="indeterminate", length=380)
        pb.pack(pady=5)
        pb.start(10)

        lbl_sub = ttk.Label(
            frame,
            text="Aguarde um momento... O sistema está processando os dados.",
            font=("Arial", 8, "italic"),
            foreground="#555555",
            justify="center",
        )
        lbl_sub.pack(pady=(5, 0))

        resultado = [None]
        erro = [None]

        def _worker():
            try:
                resultado[0] = target_fn()
            except Exception as e:
                erro[0] = e
            finally:
                def _finalizar():
                    try:
                        pb.stop()
                        top.grab_release()
                        top.destroy()
                    except Exception:
                        pass

                    if erro[0] is not None:
                        if on_erro:
                            on_erro(erro[0])
                        else:
                            messagebox.showerror("Erro", f"Ocorreu um erro durante a operação:\n{erro[0]}")
                            self.lbl_status.config(text=f"Erro: {erro[0]}")
                    else:
                        if on_sucesso:
                            on_sucesso(resultado[0])

                self.root.after(0, _finalizar)

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def ordenar_treeview(self, tree: ttk.Treeview, col: str, reverse: bool = False) -> None:
        """Ordena os itens da Treeview numericamente ou alfabeticamente ao clicar no cabeçalho.

        Permite ordenação bi-direcional inteligente (crescente/decrescente), convertendo
        valores numéricos (ex: totais, pendências) para ordenação matemática e strings para
        ordenação lexicográfica em caixa baixa.

        Args:
            tree: Instância do widget `ttk.Treeview`.
            col: Identificador da coluna clicada.
            reverse: Se True, inverte a ordem de classificação.
        """
        itens = [(tree.set(k, col), k) for k in tree.get_children("")]

        def _converter_chave(val: str) -> Any:
            v = str(val).strip()
            try:
                v_num = v.replace(".", "").replace(",", ".")
                return (0, float(v_num))
            except ValueError:
                pass
            return (1, v.lower())

        itens.sort(key=lambda t: _converter_chave(t[0]), reverse=reverse)

        for idx, (_, k) in enumerate(itens):
            tree.move(k, "", idx)

        # Inverte a ordem para o próximo clique do usuário
        tree.heading(col, command=lambda: self.ordenar_treeview(tree, col, not reverse))

    # =========================================================================
    # ABA 1: BUSCA DE HISTÓRICO DE FISCAIS
    # =========================================================================
    def _construir_aba_historico(self) -> None:
        """Constrói os componentes da Aba 1 (Filtros, Botões de Importação e Tabela de Atendimentos)."""
        top_frame = ttk.LabelFrame(self.tab_historico, text=" Pesquisa de Atendimentos Anteriores ", padding=10)
        top_frame.pack(fill="x", padx=10, pady=10)

        # Campo de busca por CNPJ ou Razão Social
        ttk.Label(top_frame, text="CNPJ ou Razão Social:", font=("Arial", 9)).pack(side="left", padx=(0, 4))
        self.ent_busca_hist = ttk.Entry(top_frame, width=26, font=("Arial", 9))
        self.ent_busca_hist.pack(side="left", padx=4)
        self.ent_busca_hist.bind("<Return>", lambda e: self.pesquisar_historico())

        # Dropdown para filtrar por nome do Fiscal
        ttk.Label(top_frame, text="Fiscal:", font=("Arial", 9)).pack(side="left", padx=(10, 4))
        self.cb_filtro_fiscal = ttk.Combobox(top_frame, state="readonly", width=18, font=("Arial", 9))
        self.cb_filtro_fiscal.pack(side="left", padx=4)
        self.cb_filtro_fiscal.bind("<<ComboboxSelected>>", lambda e: self.pesquisar_historico())

        btn_buscar = ttk.Button(top_frame, text="Pesquisar", command=self.pesquisar_historico)
        btn_buscar.pack(side="left", padx=4)

        btn_limpar = ttk.Button(top_frame, text="Limpar", command=self._limpar_busca_hist)
        btn_limpar.pack(side="left", padx=4)

        # Painel de Importação de Planilhas Antigas
        import_frame = ttk.LabelFrame(
            self.tab_historico,
            text=" Importação de Planilhas Anteriores e Aprendizado de Fiscais ",
            padding=10,
        )
        import_frame.pack(fill="x", padx=10, pady=(0, 10))

        btn_importar_pasta = ttk.Button(
            import_frame,
            text="📁 Importar da Pasta 'Planilhas anteriores'",
            command=self.importar_pasta_planilhas,
        )
        btn_importar_pasta.pack(side="left", padx=5)

        btn_importar_arq = ttk.Button(
            import_frame,
            text="📥 Selecionar Planilha Excel (.xlsx)...",
            command=self.importar_arquivo_excel,
        )
        btn_importar_arq.pack(side="left", padx=5)

        self.lbl_info_import = ttk.Label(
            import_frame,
            text="Carregando estatísticas...",
            font=("Arial", 9, "italic"),
        )
        self.lbl_info_import.pack(side="right", padx=10)

        # Tabela de Resultados
        tabela_frame = ttk.Frame(self.tab_historico)
        tabela_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        colunas = ("cnpj", "fiscal", "data", "origem", "status")
        self.tree_hist = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse")

        rotulos_hist = {
            "cnpj": ("CNPJ", 170, "center"),
            "fiscal": ("Fiscal Atribuído", 180, "w"),
            "data": ("Data Atendimento", 130, "center"),
            "origem": ("Origem / Empresa", 320, "w"),
            "status": ("Localização / Status", 140, "center"),
        }
        for c, (txt, w, anch) in rotulos_hist.items():
            self.tree_hist.heading(
                c,
                text=txt,
                command=lambda col=c: self.ordenar_treeview(self.tree_hist, col, False),
            )
            self.tree_hist.column(c, width=w, anchor=anch)

        scroll_y = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tree_hist.yview)
        self.tree_hist.configure(yscrollcommand=scroll_y.set)
        self.tree_hist.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

    def atualizar_combobox_fiscais(self) -> None:
        """Preenche o dropdown de fiscais com a lista de fiscais cadastrados."""
        try:
            fiscais = listar_fiscais_cadastrados()
            opcoes = ["Todos os Fiscais"] + fiscais
            self.cb_filtro_fiscal["values"] = opcoes
            atual = self.cb_filtro_fiscal.get()
            if not atual or atual not in opcoes:
                self.cb_filtro_fiscal.current(0)
        except Exception:
            pass

    def carregar_ultimos_fiscais(self, limit: int = 50) -> None:
        """Exibe automaticamente os atendimentos mais recentes sem necessidade de pesquisa."""
        fisc = self.cb_filtro_fiscal.get().strip() if hasattr(self, "cb_filtro_fiscal") else ""
        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)

        resultados = buscar_historico_fiscais(termo_busca="", fiscal_filtro=fisc, limit=limit)
        for r in resultados:
            self.tree_hist.insert(
                "",
                "end",
                values=(
                    r["cnpj_formatado"],
                    r["fiscal"],
                    r["data"],
                    r["origem"],
                    r["status"],
                ),
            )
        self.lbl_status.config(text=f"Exibindo os {len(resultados)} atendimentos mais recentes gravados no sistema.")

    def _limpar_busca_hist(self) -> None:
        """Limpa filtros e campos de pesquisa do histórico."""
        self.ent_busca_hist.delete(0, tk.END)
        if hasattr(self, "cb_filtro_fiscal") and self.cb_filtro_fiscal["values"]:
            self.cb_filtro_fiscal.current(0)
        self.carregar_ultimos_fiscais()
        self.lbl_status.config(text="Filtros limpos. Exibindo atendimentos recentes.")

    def pesquisar_historico(self) -> None:
        """Executa a busca de atendimentos e atualiza a Treeview."""
        termo = self.ent_busca_hist.get().strip()
        fisc = self.cb_filtro_fiscal.get().strip() if hasattr(self, "cb_filtro_fiscal") else ""

        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)

        resultados = buscar_historico_fiscais(termo_busca=termo, fiscal_filtro=fisc, limit=100)
        if not resultados:
            self.lbl_status.config(text="Nenhum atendimento encontrado para os filtros informados.")
            messagebox.showinfo("Pesquisa", "Nenhum atendimento localizado para os critérios informados.")
            return

        for r in resultados:
            self.tree_hist.insert(
                "",
                "end",
                values=(
                    r["cnpj_formatado"],
                    r["fiscal"],
                    r["data"],
                    r["origem"],
                    r["status"],
                ),
            )
        self.lbl_status.config(text=f"{len(resultados)} atendimento(s) encontrado(s).")

    def importar_pasta_planilhas(self) -> None:
        """Importa todas as planilhas da pasta 'Planilhas anteriores/' em thread separada."""
        def _executar():
            return escanear_e_importar_planilhas_antigas(forcar=True)

        def _on_sucesso(res):
            if res["total_arquivos"] == 0:
                messagebox.showinfo(
                    "Importação de Histórico",
                    "Nenhuma planilha (.xlsx ou .xls) foi encontrada na pasta 'Planilhas anteriores/'.\n\n"
                    "Coloque seus arquivos de anos/meses anteriores nessa pasta e clique novamente neste botão.",
                )
                self.lbl_status.config(text="Nenhuma planilha encontrada em 'Planilhas anteriores/'.")
                return

            msg = (
                f"Importação concluída com sucesso!\n\n"
                f"• Planilhas processadas: {res['total_arquivos']}\n"
                f"• Protocolos adicionados ao sistema: {res['total_protocolos']}\n"
                f"• Atendimentos de fiscais aprendidos: {res['total_fiscais']}\n"
                f"• Correlações de CEP mapeadas: {res['total_ceps']}\n"
                f"• Meses de competência identificados: {len(res['meses'])}\n\n"
                f"O banco de dados e as abas do sistema foram atualizados com sucesso."
            )
            messagebox.showinfo("Sucesso na Importação", msg)
            self.atualizar_tudo()

        self._executar_com_loading(
            target_fn=_executar,
            titulo="Importando Planilhas Antigas",
            mensagem="Processando e consolidando planilhas da pasta 'Planilhas anteriores'...",
            on_sucesso=_on_sucesso,
        )

    def importar_arquivo_excel(self) -> None:
        """Abre seletor para o usuário escolher uma planilha Excel e a importa sem travar a interface."""
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha de Histórico",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return

        def _executar():
            return importar_planilha_historico(caminho, forcar=True)

        def _on_sucesso(res):
            if res.get("status") == "ERRO_LEITURA":
                messagebox.showerror("Erro de Leitura", f"Falha ao ler o arquivo Excel:\n{res.get('mensagem')}")
                return

            msg = (
                f"Planilha '{res['arquivo']}' importada com êxito!\n\n"
                f"• Protocolos adicionados ao sistema: {res['protocolos_inseridos']}\n"
                f"• Atendimentos de fiscais aprendidos: {res['fiscais_inseridos']}\n"
                f"• Correlações de CEP mapeadas: {res['ceps_inseridos']}\n"
                f"• Meses identificados: {', '.join(res.get('meses_identificados', [])) or 'Nenhum'}\n\n"
                f"O sistema já memorizou os dados e está pronto para uso."
            )
            messagebox.showinfo("Sucesso", msg)
            self.atualizar_tudo()

        self._executar_com_loading(
            target_fn=_executar,
            titulo="Importando Planilha",
            mensagem=f"Lendo e consolidando dados de '{Path(caminho).name}'...",
            on_sucesso=_on_sucesso,
        )

    # =========================================================================
    # ABA 2: GESTÃO DE CEP X ÁREA
    # =========================================================================
    def _construir_aba_cep(self) -> None:
        """Constrói os componentes da Aba 2: Tabela de CEPs Mapeados, Formulário e Tabela de Pendências."""
        paned = ttk.PanedWindow(self.tab_cep, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        frame_esq = ttk.Frame(paned, padding=5)
        frame_dir = ttk.Frame(paned, padding=5)
        paned.add(frame_esq, weight=3)
        paned.add(frame_dir, weight=2)

        # --- Frame Esquerdo: CEPs Mapeados ---
        lbl_frame_map = ttk.LabelFrame(frame_esq, text=" CEPs Mapeados (cep_areas.db) ", padding=5)
        lbl_frame_map.pack(fill="both", expand=True)

        busca_cep_frame = ttk.Frame(lbl_frame_map)
        busca_cep_frame.pack(fill="x", pady=5)

        ttk.Label(busca_cep_frame, text="CEP/Termo:").pack(side="left", padx=3)
        self.ent_filtro_cep = ttk.Entry(busca_cep_frame, width=13)
        self.ent_filtro_cep.pack(side="left", padx=3)
        self.ent_filtro_cep.bind("<KeyRelease>", lambda e: self.carregar_ceps_mapeados())

        # Dropdowns: Filtrar por Área e Bairro
        ttk.Label(busca_cep_frame, text="Área:").pack(side="left", padx=(8, 3))
        self.cb_filtro_area = ttk.Combobox(busca_cep_frame, state="readonly", width=14)
        self.cb_filtro_area.pack(side="left", padx=3)
        self.cb_filtro_area.bind("<<ComboboxSelected>>", lambda e: self.carregar_ceps_mapeados())

        ttk.Label(busca_cep_frame, text="Bairro:").pack(side="left", padx=(8, 3))
        self.cb_filtro_bairro = ttk.Combobox(busca_cep_frame, state="readonly", width=16)
        self.cb_filtro_bairro.pack(side="left", padx=3)
        self.cb_filtro_bairro.bind("<<ComboboxSelected>>", lambda e: self.carregar_ceps_mapeados())

        btn_limpar_filtros_cep = ttk.Button(busca_cep_frame, text="Limpar", command=self._limpar_filtros_cep)
        btn_limpar_filtros_cep.pack(side="left", padx=4)

        cols_cep = ("cep", "area", "bairro", "origem")
        self.tree_cep = ttk.Treeview(lbl_frame_map, columns=cols_cep, show="headings", selectmode="browse")

        rotulos_cep = {
            "cep": ("CEP", 95, "center"),
            "area": ("Área", 120, "w"),
            "bairro": ("Bairro Referência", 160, "w"),
            "origem": ("Origem", 110, "center"),
        }
        for c, (txt, w, anch) in rotulos_cep.items():
            self.tree_cep.heading(
                c,
                text=txt,
                command=lambda col=c: self.ordenar_treeview(self.tree_cep, col, False),
            )
            self.tree_cep.column(c, width=w, anchor=anch)

        scroll_cep = ttk.Scrollbar(lbl_frame_map, orient="vertical", command=self.tree_cep.yview)
        self.tree_cep.configure(yscrollcommand=scroll_cep.set)
        self.tree_cep.pack(side="left", fill="both", expand=True)
        scroll_cep.pack(side="right", fill="y")
        self.tree_cep.bind("<<TreeviewSelect>>", self._ao_selecionar_cep_mapeado)

        # --- Frame Direito: Formulário e Pendências ---
        form_frame = ttk.LabelFrame(frame_dir, text=" Cadastrar / Editar Mapeamento ", padding=10)
        form_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(form_frame, text="CEP:").grid(row=0, column=0, sticky="w", pady=3)
        self.ent_form_cep = ttk.Entry(form_frame, width=18)
        self.ent_form_cep.grid(row=0, column=1, sticky="w", pady=3)

        ttk.Label(form_frame, text="Área:").grid(row=1, column=0, sticky="w", pady=3)
        self.cb_form_area = ttk.Combobox(
            form_frame,
            values=self._obter_areas_disponiveis(),
            width=16,
        )
        self.cb_form_area.grid(row=1, column=1, sticky="w", pady=3)

        ttk.Label(form_frame, text="Bairro:").grid(row=2, column=0, sticky="w", pady=3)
        self.ent_form_bairro = ttk.Entry(form_frame, width=25)
        self.ent_form_bairro.grid(row=2, column=1, sticky="w", pady=3)

        botoes_form = ttk.Frame(form_frame)
        botoes_form.grid(row=3, column=0, columnspan=2, pady=10)

        btn_salvar_cep = ttk.Button(botoes_form, text="Salvar Mapeamento", command=self.salvar_cep_form)
        btn_salvar_cep.pack(side="left", padx=5)

        btn_excluir_cep = ttk.Button(botoes_form, text="Excluir", command=self.excluir_cep_form)
        btn_excluir_cep.pack(side="left", padx=5)

        btn_limpar_form = ttk.Button(botoes_form, text="Limpar", command=self._limpar_form_cep)
        btn_limpar_form.pack(side="left", padx=5)

        # Tabela de CEPs Pendentes
        pend_frame = ttk.LabelFrame(frame_dir, text=" CEPs Pendentes de Mapeamento ", padding=5)
        pend_frame.pack(fill="both", expand=True)

        cols_pend = ("cep", "ocorrencias", "bairro")
        self.tree_pend = ttk.Treeview(pend_frame, columns=cols_pend, show="headings", selectmode="browse")

        rotulos_pend = {
            "cep": ("CEP", 95, "center"),
            "ocorrencias": ("Vezes", 55, "center"),
            "bairro": ("Bairro", 140, "w"),
        }
        for c, (txt, w, anch) in rotulos_pend.items():
            self.tree_pend.heading(
                c,
                text=txt,
                command=lambda col=c: self.ordenar_treeview(self.tree_pend, col, False),
            )
            self.tree_pend.column(c, width=w, anchor=anch)

        scroll_pend = ttk.Scrollbar(pend_frame, orient="vertical", command=self.tree_pend.yview)
        self.tree_pend.configure(yscrollcommand=scroll_pend.set)
        self.tree_pend.pack(side="left", fill="both", expand=True)
        scroll_pend.pack(side="right", fill="y")
        self.tree_pend.bind("<<TreeviewSelect>>", self._ao_selecionar_cep_pendente)

    def _obter_areas_disponiveis(self) -> list[str]:
        """Obtém lista consolidada de Áreas conhecidas."""
        padrao = [
            "CENTRAL", "HB", "BOSQUE", "REPRESA", "CIDADE DA CRIANÇA",
            "CEU", "PINHEIRINHO", "VILA TONINHO", "SCHMITT", "TALHADO"
        ]
        try:
            from src.core.database import get_triagem_db
            conn = get_triagem_db()
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT area FROM bairros_areas WHERE area IS NOT NULL AND area != '';")
            db_areas = [r[0].strip().upper() for r in cur.fetchall()]
            conn.close()
            return sorted(list(set(padrao + db_areas)))
        except Exception:
            return sorted(padrao)

    def atualizar_comboboxes_cep(self) -> None:
        """Atualiza as opções dos dropdowns de Área e Bairro na aba de CEPs."""
        try:
            areas = listar_areas_mapeadas()
            opcoes_area = ["Todas as Áreas"] + areas
            self.cb_filtro_area["values"] = opcoes_area
            if not self.cb_filtro_area.get() or self.cb_filtro_area.get() not in opcoes_area:
                self.cb_filtro_area.current(0)

            bairros = listar_bairros_mapeados()
            opcoes_bairro = ["Todos os Bairros"] + bairros
            self.cb_filtro_bairro["values"] = opcoes_bairro
            if not self.cb_filtro_bairro.get() or self.cb_filtro_bairro.get() not in opcoes_bairro:
                self.cb_filtro_bairro.current(0)

            self.cb_form_area["values"] = self._obter_areas_disponiveis()
        except Exception:
            pass

    def _limpar_filtros_cep(self) -> None:
        """Limpa filtros e campos de pesquisa de CEP."""
        self.ent_filtro_cep.delete(0, tk.END)
        if hasattr(self, "cb_filtro_area") and self.cb_filtro_area["values"]:
            self.cb_filtro_area.current(0)
        if hasattr(self, "cb_filtro_bairro") and self.cb_filtro_bairro["values"]:
            self.cb_filtro_bairro.current(0)
        self.carregar_ceps_mapeados()

    def _ao_selecionar_cep_mapeado(self, event: Any) -> None:
        """Preenche o formulário com o CEP mapeado selecionado na tabela."""
        sel = self.tree_cep.selection()
        if not sel:
            return
        vals = self.tree_cep.item(sel[0], "values")
        if vals:
            self.ent_form_cep.delete(0, tk.END)
            self.ent_form_cep.insert(0, vals[0])
            self.cb_form_area.set(vals[1])
            self.ent_form_bairro.delete(0, tk.END)
            self.ent_form_bairro.insert(0, vals[2])

    def _ao_selecionar_cep_pendente(self, event: Any) -> None:
        """Preenche o formulário com o CEP pendente selecionado na tabela."""
        sel = self.tree_pend.selection()
        if not sel:
            return
        vals = self.tree_pend.item(sel[0], "values")
        if vals:
            self.ent_form_cep.delete(0, tk.END)
            self.ent_form_cep.insert(0, vals[0])
            self.ent_form_bairro.delete(0, tk.END)
            self.ent_form_bairro.insert(0, vals[2])

    def _limpar_form_cep(self) -> None:
        """Limpa os campos do formulário de CEP."""
        self.ent_form_cep.delete(0, tk.END)
        self.cb_form_area.set("")
        self.ent_form_bairro.delete(0, tk.END)

    def carregar_ceps_mapeados(self) -> None:
        """Carrega a tabela de CEPs mapeados respeitando os filtros ativos."""
        termo = self.ent_filtro_cep.get().strip()
        area_f = self.cb_filtro_area.get().strip() if hasattr(self, "cb_filtro_area") else ""
        bairro_f = self.cb_filtro_bairro.get().strip() if hasattr(self, "cb_filtro_bairro") else ""

        for item in self.tree_cep.get_children():
            self.tree_cep.delete(item)

        ceps = listar_ceps_mapeados(
            filtro=termo,
            area_filtro=area_f,
            bairro_filtro=bairro_f,
        )
        for c in ceps:
            self.tree_cep.insert(
                "",
                "end",
                values=(
                    c["cep_formatado"] or c["cep_limpo"],
                    c["area"],
                    c["bairro_referencia"] or "",
                    c["origem"] or "",
                ),
            )

    def carregar_ceps_pendentes(self) -> None:
        """Carrega a tabela de CEPs pendentes."""
        for item in self.tree_pend.get_children():
            self.tree_pend.delete(item)

        pendentes = listar_ceps_pendentes()
        for p in pendentes:
            self.tree_pend.insert(
                "",
                "end",
                values=(
                    p["cep_formatado"] or p["cep_limpo"],
                    p["ocorrencias"],
                    p["bairro_referencia"] or "",
                ),
            )

    def salvar_cep_form(self) -> None:
        """Salva ou atualiza a correlação de CEP x Área preenchida no formulário."""
        cep_val = self.ent_form_cep.get().strip()
        area_val = self.cb_form_area.get().strip().upper()
        bairro_val = self.ent_form_bairro.get().strip()

        cep_limpo = normalizar_cep(cep_val)
        if len(cep_limpo) != 8:
            messagebox.showerror("Erro", "O CEP deve possuir 8 dígitos válidos.")
            return

        if not area_val:
            messagebox.showerror("Erro", "Selecione ou digite uma Área.")
            return

        try:
            registrar_area_cep(
                cep=cep_limpo,
                area=area_val,
                bairro=bairro_val,
                origem="MANUAL",
            )
            from src.core.cep_service import backfill_area_backlog_por_cep
            qtd_up = backfill_area_backlog_por_cep(cep_limpo, area_val)
            msg_ok = f"CEP {formatar_cep(cep_limpo)} vinculado à Área '{area_val}'."
            if qtd_up > 0:
                msg_ok += f"\n({qtd_up} protocolo(s) no backlog atualizado(s))"

            messagebox.showinfo("Sucesso", msg_ok)
            self._limpar_form_cep()
            self.atualizar_comboboxes_cep()
            self.carregar_ceps_mapeados()
            self.carregar_ceps_pendentes()
            self.lbl_status.config(text=f"CEP {formatar_cep(cep_limpo)} salvo com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", str(e))

    def excluir_cep_form(self) -> None:
        """Exclui a correlação de CEP informada no formulário."""
        cep_val = self.ent_form_cep.get().strip()
        cep_limpo = normalizar_cep(cep_val)
        if not cep_limpo:
            messagebox.showwarning("Aviso", "Selecione um CEP para excluir.")
            return

        if messagebox.askyesno("Confirmar Exclusão", f"Deseja excluir o mapeamento do CEP {formatar_cep(cep_limpo)}?"):
            remover_cep_mapeado(cep_limpo)
            self._limpar_form_cep()
            self.atualizar_comboboxes_cep()
            self.carregar_ceps_mapeados()
            self.lbl_status.config(text=f"CEP {formatar_cep(cep_limpo)} excluído.")

    # =========================================================================
    # ABA 3: ARQUIVAMENTO DE COMPETÊNCIAS
    # =========================================================================
    def _construir_aba_arquivamento(self) -> None:
        """Constrói os componentes da Aba 3 (KPIs, Tabela de Competências e Botão de Aplicação)."""
        info_frame = ttk.LabelFrame(self.tab_arquivamento, text=" Controle de Visibilidade e Arquivamento por Mês ", padding=10)
        info_frame.pack(fill="x", padx=10, pady=(10, 5))

        lbl_desc = ttk.Label(
            info_frame,
            text=(
                "• Gerenciamento por Mês de Competência: defina quais meses devem constar na planilha de trabalho ativa.\n"
                "• Ativo (Visível): os protocolos deste mês aparecem na planilha de trabalho atual.\n"
                "• Arquivado (Oculto): os protocolos deste mês são ocultados da planilha de trabalho ativa (dados permanecem salvos no SQLite).\n"
                "• Aplicação: ao clicar em 'Aplicar Alterações', a planilha anterior é guardada em 'Planilhas anteriores/' "
                "e uma nova planilha é gerada com os meses ativos e novos pedidos do CSV."
            ),
            font=("Arial", 9),
            justify="left",
        )
        lbl_desc.pack(fill="x", padx=5, pady=2)

        # Cards de Estatísticas / KPIs
        kpi_frame = ttk.LabelFrame(self.tab_arquivamento, text=" Panorama Geral ", padding=10)
        kpi_frame.pack(fill="x", padx=10, pady=5)

        self.lbl_kpi_total = ttk.Label(kpi_frame, text="Backlog Total: 0", font=("Arial", 10, "bold"))
        self.lbl_kpi_total.pack(side="left", padx=12)

        self.lbl_kpi_visiveis = ttk.Label(kpi_frame, text="Ativos na Planilha: 0", font=("Arial", 10, "bold"), foreground="#27AE60")
        self.lbl_kpi_visiveis.pack(side="left", padx=12)

        self.lbl_kpi_pend = ttk.Label(kpi_frame, text="Pendentes: 0", font=("Arial", 10))
        self.lbl_kpi_pend.pack(side="left", padx=12)

        self.lbl_kpi_fin = ttk.Label(kpi_frame, text="Finalizados: 0", font=("Arial", 10))
        self.lbl_kpi_fin.pack(side="left", padx=12)

        self.lbl_kpi_hist = ttk.Label(kpi_frame, text="Histórico: 0", font=("Arial", 10), foreground="#2980B9")
        self.lbl_kpi_hist.pack(side="left", padx=12)

        self.lbl_kpi_meses = ttk.Label(kpi_frame, text="Meses: 0 Ativos / 0 Arquivados", font=("Arial", 10, "bold"), foreground="#8E44AD")
        self.lbl_kpi_meses.pack(side="left", padx=12)

        # Tabela de Meses de Competência
        tabela_frame = ttk.LabelFrame(self.tab_arquivamento, text=" Meses de Competência Cadastrados ", padding=10)
        tabela_frame.pack(fill="both", expand=True, padx=10, pady=5)

        cols_lote = ("competencia", "status", "total", "pendentes", "distribuidos", "finalizados")
        self.tree_lotes = ttk.Treeview(tabela_frame, columns=cols_lote, show="headings", selectmode="browse")

        rotulos_lote = {
            "competencia": ("Mês / Competência", 140, "center"),
            "status": ("Status na Planilha de Trabalho", 240, "center"),
            "total": ("Total Protocolos", 120, "center"),
            "pendentes": ("Pendentes", 110, "center"),
            "distribuidos": ("Distribuídos", 110, "center"),
            "finalizados": ("Finalizados", 110, "center"),
        }
        for c, (txt, w, anch) in rotulos_lote.items():
            self.tree_lotes.heading(
                c,
                text=txt,
                command=lambda col=c: self.ordenar_treeview(self.tree_lotes, col, False),
            )
            self.tree_lotes.column(c, width=w, anchor=anch)

        scroll_lotes = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tree_lotes.yview)
        self.tree_lotes.configure(yscrollcommand=scroll_lotes.set)
        self.tree_lotes.pack(side="left", fill="both", expand=True)
        scroll_lotes.pack(side="right", fill="y")
        self.tree_lotes.bind("<Double-1>", lambda e: self.alternar_status_selecionado())

        # Botões de Ação
        botoes_frame = ttk.Frame(self.tab_arquivamento)
        botoes_frame.pack(fill="x", padx=10, pady=8)

        btn_alternar = ttk.Button(
            botoes_frame,
            text="🔄 Alternar Status (Duplo Clique)",
            command=self.alternar_status_selecionado,
        )
        btn_alternar.pack(side="left", padx=5)

        btn_arquivar = ttk.Button(
            botoes_frame,
            text="📦 Arquivar Mês (Ocultar)",
            command=self.arquivar_selecionado,
        )
        btn_arquivar.pack(side="left", padx=5)

        btn_reativar = ttk.Button(
            botoes_frame,
            text="👁️ Reativar Mês (Exibir)",
            command=self.reativar_selecionado,
        )
        btn_reativar.pack(side="left", padx=5)

        btn_aplicar = tk.Button(
            botoes_frame,
            text="💾 Aplicar Alterações na Planilha de Trabalho",
            bg="#27AE60",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=18,
            pady=6,
            command=self.aplicar_alteracoes_planilha,
        )
        btn_aplicar.pack(side="left", padx=15)

        btn_atualizar = ttk.Button(botoes_frame, text="Atualizar Lista", command=self.carregar_lotes_arquivamento)
        btn_atualizar.pack(side="right", padx=5)

    def carregar_lotes_arquivamento(self) -> None:
        """Carrega e calcula KPIs do backlog e preenche a tabela de competências."""
        stats = obter_estatisticas_gerais()
        self.lbl_kpi_total.config(text=f"Backlog Total: {stats['total_backlog']}")
        self.lbl_kpi_visiveis.config(text=f"Ativos na Planilha: {stats['protocolos_ativos_visiveis']}")
        self.lbl_kpi_pend.config(text=f"Pendentes: {stats['total_pendentes']}")
        self.lbl_kpi_fin.config(text=f"Finalizados: {stats['total_finalizados']}")
        self.lbl_kpi_hist.config(text=f"Histórico: {stats['total_historico']}")
        self.lbl_kpi_meses.config(text=f"Meses: {stats['meses_ativos']} Ativos / {stats['meses_arquivados']} Arquivados")

        for item in self.tree_lotes.get_children():
            self.tree_lotes.delete(item)

        lotes = listar_competencias_status()
        for lote in lotes:
            status_txt = "📦 Arquivado (Oculto)" if lote["arquivado"] else "✅ Ativo (Visível na Planilha)"
            self.tree_lotes.insert(
                "",
                "end",
                values=(
                    lote["mes_competencia"],
                    status_txt,
                    lote["total"],
                    lote["pendentes"],
                    lote["distribuidos"],
                    lote["finalizados"],
                ),
            )

    def alternar_status_selecionado(self) -> None:
        """Alterna o status do mês selecionado entre Ativo e Arquivado."""
        sel = self.tree_lotes.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um mês de competência na tabela.")
            return

        vals = self.tree_lotes.item(sel[0], "values")
        competencia = vals[0]
        novo_arquivado = alternar_status_competencia(competencia)
        status_novo = "Arquivado (Oculto)" if novo_arquivado else "Ativo (Visível)"
        self.carregar_lotes_arquivamento()
        self.lbl_status.config(
            text=f"Mês {competencia} marcado como {status_novo}. Clique em 'Aplicar Alterações' para atualizar a planilha."
        )

    def arquivar_selecionado(self) -> None:
        """Marca o mês selecionado como Arquivado."""
        sel = self.tree_lotes.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um mês de competência na tabela.")
            return

        vals = self.tree_lotes.item(sel[0], "values")
        competencia = vals[0]
        definir_status_competencia(competencia, arquivado=True)
        self.carregar_lotes_arquivamento()
        self.lbl_status.config(
            text=f"Mês {competencia} marcado como Arquivado (Oculto). Clique em 'Aplicar Alterações' para atualizar a planilha."
        )

    def reativar_selecionado(self) -> None:
        """Reativa o mês selecionado (marca como Ativo / Visível)."""
        sel = self.tree_lotes.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um mês de competência na tabela.")
            return

        vals = self.tree_lotes.item(sel[0], "values")
        competencia = vals[0]
        definir_status_competencia(competencia, arquivado=False)
        self.carregar_lotes_arquivamento()
        self.lbl_status.config(
            text=f"Mês {competencia} reativado (Visível). Clique em 'Aplicar Alterações' para atualizar a planilha."
        )

    def aplicar_alteracoes_planilha(self) -> None:
        """Aplica reorganização de competências, gera backup e recria a planilha ativa."""
        if not messagebox.askyesno(
            "Aplicar Alterações e Atualizar Planilha",
            "Deseja atualizar a planilha de trabalho ativa?\n\n"
            "• A planilha que estava sendo usada será salva como cópia de segurança em 'Planilhas anteriores/'.\n"
            "• Uma nova 'Planilha_Controle_VRE_Ativa.xlsx' será gerada com os meses selecionados e eventuais novos pedidos do CSV.\n\n"
            "Continuar?",
        ):
            return

        def _executar():
            return aplicar_reorganizacao_planilha()

        def _on_sucesso(res):
            messagebox.showinfo("Sucesso", res["mensagem"])
            self.atualizar_tudo()
            self.lbl_status.config(text=f"Planilha de trabalho atualizada ({res['linhas_geradas']} protocolos).")

        self._executar_com_loading(
            target_fn=_executar,
            titulo="Atualizando Planilha de Trabalho",
            mensagem="Gerando nova planilha de trabalho oficial com os meses ativos...",
            on_sucesso=_on_sucesso,
        )

    def executar_arquivamento(self) -> None:
        """Compatibilidade: executa arquivamento/alternância do mês selecionado."""
        self.alternar_status_selecionado()

    def atualizar_tudo(self) -> None:
        """Recarrega os dados em todas as abas e componentes da interface."""
        self.atualizar_combobox_fiscais()
        self.carregar_ultimos_fiscais()
        self.atualizar_comboboxes_cep()
        self.carregar_ceps_mapeados()
        self.carregar_ceps_pendentes()
        self.carregar_lotes_arquivamento()
        try:
            stats = obter_estatisticas_gerais()
            total_ceps = len(listar_ceps_mapeados())
            txt_info = f"Histórico: {stats.get('total_historico', 0)} atendimentos | {total_ceps} CEPs mapeados"
            if hasattr(self, "lbl_info_import"):
                self.lbl_info_import.config(text=txt_info)
        except Exception:
            pass


def executar_modo_cli(args: argparse.Namespace) -> None:
    """Modo texto/CLI para administração quando GUI não puder ser aberta.

    Args:
        args: Argumentos parseados via `argparse`.
    """
    if getattr(args, "buscar", None):
        termo = args.buscar.strip()
        print(f"[*] Pesquisando histórico para: '{termo}'...")
        res = buscar_historico_fiscais(termo)
        if not res:
            print("    [!] Nenhum registro encontrado.")
        else:
            print(f"    [✓] {len(res)} registro(s) encontrado(s):")
            for r in res:
                print(f"      - CNPJ: {r['cnpj_formatado']} | Fiscal: {r['fiscal']} | Data: {r['data']} | Status: {r['status']} | Origem: {r['origem']}")
        return

    if getattr(args, "arquivar", None):
        comp = args.arquivar.strip()
        print(f"[*] Marcando competência {comp} como arquivada (oculta da planilha)...")
        ok = definir_status_competencia(comp, arquivado=True)
        if ok:
            print(f"    [✓] Competência {comp} arquivada com sucesso.")
        else:
            print(f"    [!] Falha ao arquivar competência {comp}.")
        return

    if getattr(args, "reativar", None):
        comp = args.reativar.strip()
        print(f"[*] Reativando competência {comp} (visível na planilha)...")
        ok = definir_status_competencia(comp, arquivado=False)
        if ok:
            print(f"    [✓] Competência {comp} reativada com sucesso.")
        else:
            print(f"    [!] Falha ao reativar competência {comp}.")
        return

    if getattr(args, "aplicar", False):
        print("[*] Aplicando alterações e gerando nova planilha de trabalho...")
        res = aplicar_reorganizacao_planilha()
        print(f"    [✓] {res['mensagem']}")
        return

    if getattr(args, "mapear_cep", None):
        cep = args.mapear_cep.strip()
        area = args.area.strip().upper() if getattr(args, "area", None) else ""
        if not area:
            print("    [ERRO] É obrigatório fornecer a --area ao mapear um CEP.")
            sys.exit(1)
        bairro = args.bairro.strip() if getattr(args, "bairro", None) else ""
        print(f"[*] Mapeando CEP {cep} para Área '{area}'...")
        try:
            registrar_area_cep(cep, area, bairro=bairro, origem="MANUAL")
            from src.core.cep_service import backfill_area_backlog_por_cep
            qtd_up = backfill_area_backlog_por_cep(normalizar_cep(cep), area)
            print(f"    [✓] CEP mapeado com sucesso! ({qtd_up} protocolos atualizados no backlog)")
        except Exception as e:
            print(f"    [!] Erro: {e}")
        return

    if getattr(args, "listar_pendentes", False):
        pendentes = listar_ceps_pendentes(limit=50)
        print(f"[*] CEPs Pendentes de Mapeamento ({len(pendentes)}):")
        for p in pendentes:
            print(f"    - CEP: {p['cep_formatado']} | Vezes: {p['ocorrencias']} | Bairro: {p['bairro_referencia'] or '-'} | Logradouro: {p['logradouro'] or '-'}")
        return

    if getattr(args, "listar_lotes", False):
        lotes = listar_competencias_status()
        print(f"[*] Meses de Competência ({len(lotes)}):")
        for l in lotes:
            st = "[ARQUIVADO (OCULTO)]" if l["arquivado"] else "[ATIVO (VISÍVEL)]"
            print(f"    - {l['mes_competencia']}: {st} | Total={l['total']}, Pend={l['pendentes']}, Fin={l['finalizados']}")
        return

    # Se nenhum comando específico foi passado, exibe o painel de status padrão
    print("=" * 60)
    print("SMAURB - PAINEL ADMINISTRATIVO VRE (MODO CLI / MANUTENÇÃO)")
    print("=" * 60)
    stats = obter_estatisticas_gerais()
    print(f"Backlog Total:        {stats['total_backlog']}")
    print(f"Protocolos na Planilha: {stats['protocolos_ativos_visiveis']}")
    print(f"  - Pendentes:        {stats['total_pendentes']}")
    print(f"  - Distribuídos:     {stats['total_distribuidos']}")
    print(f"  - Finalizados:      {stats['total_finalizados']}")
    print(f"Histórico Permanente: {stats['total_historico']}")
    print(f"Meses de Competência: {stats['meses_ativos']} Ativos / {stats['meses_arquivados']} Arquivados")
    print("-" * 60)

    lotes = listar_competencias_status()
    print("Meses de Competência:")
    for l in lotes:
        st = "[ARQUIVADO (OCULTO)]" if l["arquivado"] else "[ATIVO (VISÍVEL)]"
        print(f"  {l['mes_competencia']}: {st} | Total={l['total']}, Pend={l['pendentes']}, Fin={l['finalizados']}")

    pendentes_cep = listar_ceps_pendentes(limit=5)
    print("-" * 60)
    print(f"CEPs Pendentes de Mapeamento ({len(pendentes_cep)} principais):")
    for p in pendentes_cep:
        print(f"  CEP: {p['cep_formatado']} | Vezes: {p['ocorrencias']} | Bairro: {p['bairro_referencia']}")
    print("=" * 60)


def main() -> None:
    """Ponto de entrada do executável administrativo (config_vre.exe)."""
    parser = argparse.ArgumentParser(description="Painel Administrativo do VRE (SMAURB)")
    parser.add_argument("--cli", action="store_true", help="Executa em modo linha de comando (sem GUI)")
    parser.add_argument("--buscar", type=str, help="Pesquisa histórico por CNPJ ou Razão Social")
    parser.add_argument("--arquivar", type=str, help="Arquiva mês de competência (formato AAAA-MM)")
    parser.add_argument("--reativar", type=str, help="Reativa mês de competência (formato AAAA-MM)")
    parser.add_argument("--aplicar", action="store_true", help="Aplica reorganização e atualiza planilha ativa")
    parser.add_argument("--forcar", action="store_true", help="Compatibilidade")
    parser.add_argument("--mapear-cep", type=str, help="Cadastra ou atualiza o CEP informado")
    parser.add_argument("--area", type=str, help="Área municipal a associar ao CEP")
    parser.add_argument("--bairro", type=str, help="Bairro de referência para o CEP")
    parser.add_argument("--listar-pendentes", action="store_true", help="Lista CEPs pendentes de mapeamento")
    parser.add_argument("--listar-lotes", action="store_true", help="Lista lotes de competência e progresso")
    args = parser.parse_args()

    tem_comando_cli = any([
        args.cli,
        args.buscar,
        args.arquivar,
        args.reativar,
        args.aplicar,
        args.mapear_cep,
        args.listar_pendentes,
        args.listar_lotes,
    ])

    if tem_comando_cli or not TKINTER_DISPONIVEL:
        if not TKINTER_DISPONIVEL and not tem_comando_cli:
            print("[!] Tkinter não encontrado neste ambiente. Executando em modo CLI...")
        executar_modo_cli(args)
        return

    root = tk.Tk()
    app = PainelAdminVRE(root)
    root.mainloop()


if __name__ == "__main__":
    main()
