import customtkinter as ctk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import openpyxl
import os
import datetime

class PesquisarApp(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Pesquisar Clientes")
        # Ajuste conforme desejar. Ex.: "1050x700", "1200x800", etc.
        self.geometry("1000x700")
        self.resizable(True, True)

        self.create_pesquisar()

    def create_pesquisar(self):
        container = ctk.CTkFrame(self)
        container.pack(fill=BOTH, expand=True, padx=20, pady=10)

        # Título
        ctk.CTkLabel(
            container,
            text="PESQUISAR",
            font=("Century Gothic bold", 24),
            text_color="blue"
        ).pack(pady=10)

        # Frame de filtros de pesquisa
        filter_frame = ctk.CTkFrame(container)
        filter_frame.pack(pady=10, fill=X)

        # =========================
        # Variáveis de pesquisa
        # =========================
        self.id_value = StringVar()
        self.nivel_value = StringVar()
        self.classe_value = StringVar()
        self.rs_pf_value = StringVar()       # Razão Social ou Pessoa Física
        self.nome_fantasia_value = StringVar()
        self.endereco_value = StringVar()
        self.cnpj_cpf_value = StringVar()
        self.cnae_value = StringVar()
        self.parecer_value = StringVar()
        self.ultima_inspecao_value = StringVar()
        self.alvara_value = StringVar()
        self.vigi_risco_value = StringVar()
        self.observacao_value = StringVar()
        self.baixados_value = StringVar()
        self.excluidos_value = StringVar()

        # Carrega IDs, Classes, etc. do Excel (se necessário)
        self.carregar_opcoes()

        # ======================================
        # ORGANIZAÇÃO DOS CAMPOS (Exemplo)
        # ======================================
        # Linha 0
        ctk.CTkLabel(filter_frame, text="ID:", font=("Century Gothic", 14)).grid(row=0, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkComboBox(filter_frame, variable=self.id_value, values=self.id_options).grid(row=0, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="NÍVEL:", font=("Century Gothic", 14)).grid(row=0, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkComboBox(filter_frame, variable=self.nivel_value, values=["NIVEL 1", "NIVEL 2", "NIVEL 3"]).grid(row=0, column=3, padx=5, pady=5, sticky=W)

        # Linha 1
        ctk.CTkLabel(filter_frame, text="CLASSE:", font=("Century Gothic", 14)).grid(row=1, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkComboBox(filter_frame, variable=self.classe_value, values=self.classe_options).grid(row=1, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="RAZÃO SOCIAL/PF:", font=("Century Gothic", 14)).grid(row=1, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.rs_pf_value).grid(row=1, column=3, padx=5, pady=5, sticky=W)

        # Linha 2
        ctk.CTkLabel(filter_frame, text="NOME FANTASIA:", font=("Century Gothic", 14)).grid(row=2, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.nome_fantasia_value).grid(row=2, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="ENDEREÇO:", font=("Century Gothic", 14)).grid(row=2, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.endereco_value).grid(row=2, column=3, padx=5, pady=5, sticky=W)

        # Linha 3
        ctk.CTkLabel(filter_frame, text="CNPJ/CPF:", font=("Century Gothic", 14)).grid(row=3, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.cnpj_cpf_value).grid(row=3, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="CNAE (Principal):", font=("Century Gothic", 14)).grid(row=3, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.cnae_value).grid(row=3, column=3, padx=5, pady=5, sticky=W)

        # Linha 4
        ctk.CTkLabel(filter_frame, text="PARECER TÉCNICO:", font=("Century Gothic", 14)).grid(row=4, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.parecer_value).grid(row=4, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="ÚLTIMA INSPEÇÃO:", font=("Century Gothic", 14)).grid(row=4, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.ultima_inspecao_value).grid(row=4, column=3, padx=5, pady=5, sticky=W)

        # Linha 5
        ctk.CTkLabel(filter_frame, text="ALVARÁ:", font=("Century Gothic", 14)).grid(row=5, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.alvara_value).grid(row=5, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="VIGI-RISCO:", font=("Century Gothic", 14)).grid(row=5, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.vigi_risco_value).grid(row=5, column=3, padx=5, pady=5, sticky=W)

        # Linha 6
        ctk.CTkLabel(filter_frame, text="OBSERVAÇÕES:", font=("Century Gothic", 14)).grid(row=6, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkEntry(filter_frame, textvariable=self.observacao_value).grid(row=6, column=1, padx=5, pady=5, sticky=W)

        ctk.CTkLabel(filter_frame, text="BAIXADOS:", font=("Century Gothic", 14)).grid(row=6, column=2, padx=5, pady=5, sticky=W)
        ctk.CTkComboBox(filter_frame, variable=self.baixados_value, values=["BAIXADO", "Não", ""]).grid(row=6, column=3, padx=5, pady=5, sticky=W)

        # Linha 7
        ctk.CTkLabel(filter_frame, text="EXCLUIDOS:", font=("Century Gothic", 14)).grid(row=7, column=0, padx=5, pady=5, sticky=W)
        ctk.CTkComboBox(filter_frame, variable=self.excluidos_value, values=["EXCLUÍDO", "Não", ""]).grid(row=7, column=1, padx=5, pady=5, sticky=W)

        # ================
        # Frame de botões
        # ================
        button_frame = ctk.CTkFrame(container)
        button_frame.pack(pady=10)

        ctk.CTkButton(
            button_frame,
            text="Pesquisar",
            command=self.pesquisar,
            fg_color="green"
        ).grid(row=0, column=0, padx=10, pady=5)

        ctk.CTkButton(
            button_frame,
            text="Limpar",
            command=self.limpar,
            fg_color="gray"
        ).grid(row=0, column=1, padx=10, pady=5)

        ctk.CTkButton(
            button_frame,
            text="Imprimir",
            command=self.imprimir,
            fg_color="blue"
        ).grid(row=0, column=2, padx=10, pady=5)

        # ================
        # Label de estatística
        # ================
        self.stats_label = ctk.CTkLabel(
            container,
            text="Resultados encontrados: 0",
            font=("Century Gothic", 14)
        )
        self.stats_label.pack(pady=5)

        # ================
        # Frame de resultados (Treeview)
        # ================
        self.result_frame = ctk.CTkFrame(container)
        self.result_frame.pack(fill=BOTH, expand=True, pady=10)

        self.tree = ttk.Treeview(self.result_frame)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)

        self.scrollbar = Scrollbar(self.result_frame, orient=VERTICAL, command=self.tree.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.tree.configure(yscrollcommand=self.scrollbar.set)

    def carregar_opcoes(self):
        """
        Carrega as opções de ID e Classe a partir do Excel, 
        para popular os comboboxes. Ajuste caso deseje carregar outros campos.
        """
        try:
            workbook = openpyxl.load_workbook("clientes.xlsx")
            folha = workbook.active

            self.id_options = []
            self.classe_options = []

            for row in folha.iter_rows(min_row=2, values_only=True):
                # row[0] = ID, row[2] = CLASSE (exemplo)
                if row[0] not in (None, ""):
                    self.id_options.append(str(row[0]))
                if len(row) >= 3 and row[2] not in (None, ""):
                    self.classe_options.append(str(row[2]))

            # Remove duplicados e ordena
            self.id_options = sorted(set(self.id_options))
            self.classe_options = sorted(set(self.classe_options))

        except FileNotFoundError:
            self.id_options = []
            self.classe_options = []

    def pesquisar(self):
        """
        Lê o arquivo Excel, filtra os dados com base nos valores
        informados e exibe no Treeview.
        """
        try:
            workbook = openpyxl.load_workbook("clientes.xlsx")
            folha = workbook.active

            # Limpa resultados anteriores
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Define as colunas do Treeview com base na primeira linha do Excel
            cabecalho = [cell.value for cell in folha[1] if cell.value is not None]
            self.tree["columns"] = cabecalho
            self.tree["show"] = "headings"

            for col in cabecalho:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=150, anchor=W)

            total = 0

            # Percorre as linhas (a partir da 2ª) e filtra
            for row in folha.iter_rows(min_row=2, values_only=True):
                # Se a linha não tem colunas suficientes, pula
                if len(row) < len(cabecalho):
                    continue

                if self.filtrar(row):
                    linha_formatada = [str(valor) if valor is not None else "" for valor in row[:len(cabecalho)]]
                    self.tree.insert("", END, values=linha_formatada)
                    total += 1

            self.stats_label.configure(text=f"Resultados encontrados: {total}")

        except FileNotFoundError:
            messagebox.showerror("Erro", "O arquivo 'clientes.xlsx' não foi encontrado.")

    def filtrar(self, row):
        """
        Aplica os critérios de filtro em cada linha do Excel.
        Ajuste a ordem dos índices de coluna conforme sua planilha.
        Supõe-se a seguinte ordem:
          0: ID
          1: Nível
          2: Classe
          3: Razão Social / PF
          4: Nome Fantasia
          5: Endereço
          6: CNPJ/CPF
          7: CNAE (Principal)
          8: Nº Parecer Técnico
          9: Última Inspeção
         10: Alvará
         11: Vigi-Risco
         12: Observações
         13: Baixados
         14: Excluídos
        """
        # Certifique-se de que existem pelo menos 15 colunas
        if len(row) < 15:
            return False

        # Converte None -> ""
        id_excel          = str(row[0]) if row[0] else ""
        nivel_excel       = str(row[1]) if row[1] else ""
        classe_excel      = str(row[2]) if row[2] else ""
        rs_pf_excel       = str(row[3]) if row[3] else ""
        nome_fant_excel   = str(row[4]) if row[4] else ""
        endereco_excel    = str(row[5]) if row[5] else ""
        cnpj_cpf_excel    = str(row[6]) if row[6] else ""
        cnae_excel        = str(row[7]) if row[7] else ""
        parecer_excel     = str(row[8]) if row[8] else ""
        ultima_insp_excel = row[9]  # Pode ser data ou string
        alvara_excel      = str(row[10]) if row[10] else ""
        vigi_risco_excel  = str(row[11]) if row[11] else ""
        obs_excel         = str(row[12]) if row[12] else ""
        baixado_excel     = str(row[13]) if row[13] else ""
        excluido_excel    = str(row[14]) if row[14] else ""

        # ====================
        # Aplica filtros
        # ====================
        if self.id_value.get().strip():
            if id_excel != self.id_value.get().strip():
                return False

        if self.nivel_value.get().strip():
            if nivel_excel != self.nivel_value.get().strip():
                return False

        if self.classe_value.get().strip():
            if classe_excel != self.classe_value.get().strip():
                return False

        filtro_rs_pf = self.rs_pf_value.get().strip().lower()
        if filtro_rs_pf:
            if filtro_rs_pf not in rs_pf_excel.lower():
                return False

        filtro_nome = self.nome_fantasia_value.get().strip().lower()
        if filtro_nome:
            if filtro_nome not in nome_fant_excel.lower():
                return False

        filtro_end = self.endereco_value.get().strip().lower()
        if filtro_end:
            if filtro_end not in endereco_excel.lower():
                return False

        filtro_cnpj_cpf = self.cnpj_cpf_value.get().strip().lower()
        if filtro_cnpj_cpf:
            if filtro_cnpj_cpf not in cnpj_cpf_excel.lower():
                return False

        filtro_cnae = self.cnae_value.get().strip().lower()
        if filtro_cnae:
            if filtro_cnae not in cnae_excel.lower():
                return False

        filtro_parecer = self.parecer_value.get().strip().lower()
        if filtro_parecer:
            if filtro_parecer not in parecer_excel.lower():
                return False

        filtro_inspecao = self.ultima_inspecao_value.get().strip().lower()
        if filtro_inspecao:
            # Se a coluna estiver em formato de data
            if isinstance(ultima_insp_excel, (datetime.date, datetime.datetime)):
                inspecao_str = ultima_insp_excel.strftime("%d/%m/%Y")
            else:
                inspecao_str = str(ultima_insp_excel) if ultima_insp_excel else ""

            # Se o usuário digitou 4 dígitos (ex.: ano)
            if len(filtro_inspecao) == 4 and filtro_inspecao.isdigit():
                if filtro_inspecao not in inspecao_str:
                    return False
            else:
                # Busca parcial
                if filtro_inspecao not in inspecao_str.lower():
                    return False

        filtro_alvara = self.alvara_value.get().strip().lower()
        if filtro_alvara:
            if filtro_alvara not in alvara_excel.lower():
                return False

        filtro_vigi = self.vigi_risco_value.get().strip().lower()
        if filtro_vigi:
            if filtro_vigi not in vigi_risco_excel.lower():
                return False

        filtro_obs = self.observacao_value.get().strip().lower()
        if filtro_obs:
            if filtro_obs not in obs_excel.lower():
                return False

        if self.baixados_value.get().strip():
            if baixado_excel.lower() != self.baixados_value.get().strip().lower():
                return False

        if self.excluidos_value.get().strip():
            if excluido_excel.lower() != self.excluidos_value.get().strip().lower():
                return False

        return True

    def limpar(self):
        """
        Limpa os campos de pesquisa e o Treeview.
        """
        self.id_value.set("")
        self.nivel_value.set("")
        self.classe_value.set("")
        self.rs_pf_value.set("")
        self.nome_fantasia_value.set("")
        self.endereco_value.set("")
        self.cnpj_cpf_value.set("")
        self.cnae_value.set("")
        self.parecer_value.set("")
        self.ultima_inspecao_value.set("")
        self.alvara_value.set("")
        self.vigi_risco_value.set("")
        self.observacao_value.set("")
        self.baixados_value.set("")
        self.excluidos_value.set("")

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.stats_label.configure(text="Resultados encontrados: 0")

    def imprimir(self):
        """
        Salva o conteúdo do Treeview em um arquivo .txt e
        envia para impressão (somente Windows).
        """
        try:
            arquivo = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os Arquivos", "*.*")]
            )
            if not arquivo:
                return

            with open(arquivo, "w", encoding="utf-8") as f:
                for item in self.tree.get_children():
                    linha = "\t".join(self.tree.item(item, "values"))
                    f.write(linha + "\n")

            # Impressão (funciona somente no Windows, pois usa "os.startfile")
            os.startfile(arquivo, "print")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao imprimir os resultados: {e}")
