import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, date
import os
from typing import Optional, Dict, Tuple, List, Any

NOME_PLANILHA_ESTOQUE: str = 'Estoque'
NOME_PLANILHA_EXCLUIDOS: str = 'Excluidos'
DIAS_ALERTA_AMARELO: int = 30
LISTA_CATEGORIAS: List[str] = ["Pães", "Bebidas", "Picolés", "Pinguinos", "Biscoitos", "Grãos", "Lactea", "Gelatos kg", "Tortas"]
COLUNAS_ESTOQUE: List[str] = ['Produto', 'Categoria', 'Validade', 'Lote', 'Data Recebimento']
COLUNAS_EXCLUIDOS: List[str] = ['Produto', 'Categoria', 'Validade', 'Lote','Quantidade', 'Data Recebimento', 'Data Exclusao']
TREEVIEW_COL_MAP = {
    'produto': 'Produto', 'categoria': 'Categoria', 'validade': 'Validade',
    'lote': 'Lote', 'data_recebimento': 'Data Recebimento'
}
# Coluna padrão para ordenação inicial
DEFAULT_SORT_COL = 'validade'

CAMINHO_DADOS = "dados/estoque.xlsx"
CAMINHO_PLANILHA = os.path.join(CAMINHO_DADOS, "estoque.xlsx")


def verificar_ou_criar_planilha():
    # Verifica se a pasta "dados" existe
    if not os.path.exists(CAMINHO_DADOS):
        os.makedirs(CAMINHO_DADOS)

    # Verifica se a planilha existe
    if not os.path.exists(CAMINHO_PLANILHA):
        df_vazio = pd.DataFrame(columns=COLUNAS_ESTOQUE)
        df_vazio.to_excel(CAMINHO_PLANILHA, index=False)
        print(f"Planilha criada automaticamente em: {CAMINHO_PLANILHA}")

class StockManagerApp:
    """
    Aplicação para Gerenciamento de Validade de Estoque v7.
    Inclui Filtro por Nome.
    """
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Gerenciador de Validade v7 (com Filtro)")
        self.root.geometry("1050x750") # Um pouco mais de altura para o filtro

        self.status_var = tk.StringVar()
        self.search_var = tk.StringVar() # Variavel para o campo de busca
        self.editing_index: Optional[Any] = None
        self.last_sort_col: Optional[str] = DEFAULT_SORT_COL # Ordenação inicial padrão
        self.last_sort_reverse: bool = False # Ascendente por padrão

        self.df_estoque: pd.DataFrame = pd.DataFrame(columns=COLUNAS_ESTOQUE)
        self.df_excluidos: pd.DataFrame = pd.DataFrame(columns=COLUNAS_EXCLUIDOS)

        self.load_data() # Carrega e aplica ordenação inicial
        self.create_widgets()
        self.refresh_active_stock_view() # Chamada inicial para exibir dados
        self.update_deleted_display()
        self.set_status("Pronto.")

    
    
    def _handle_datetime_conversion(self, series: pd.Series) -> pd.Series:
        return pd.to_datetime(series, errors='coerce')

    def load_data(self) -> None:
        self.set_status(f"Carregando dados de {CAMINHO_PLANILHA}")
        if os.path.exists(CAMINHO_PLANILHA):
            try:
                excel_data: Dict[str, pd.DataFrame] = pd.read_excel(CAMINHO_PLANILHA, sheet_name=None, engine='openpyxl')
                # Carrega Estoque
                if NOME_PLANILHA_ESTOQUE in excel_data:
                    self.df_estoque = excel_data[NOME_PLANILHA_ESTOQUE]; self.df_estoque = self.df_estoque.dropna(how='all') # Remove linhas totalmente vazias
                    for col in COLUNAS_ESTOQUE:
                        if col not in self.df_estoque.columns: self.df_estoque[col] = pd.NaT
                    self.df_estoque['Validade'] = self._handle_datetime_conversion(self.df_estoque['Validade'])
                    self.df_estoque['Data Recebimento'] = self._handle_datetime_conversion(self.df_estoque['Data Recebimento'])
                    self.df_estoque['Lote'] = self.df_estoque['Lote'].fillna('').astype(str)
                    self.df_estoque['Categoria'] = self.df_estoque['Categoria'].fillna('').astype(str)
                    self.df_estoque.dropna(subset=['Validade'], inplace=True)
                    self.df_estoque = self.df_estoque[COLUNAS_ESTOQUE]
                else: self.df_estoque = pd.DataFrame(columns=COLUNAS_ESTOQUE)
                # Carrega Excluídos
                if NOME_PLANILHA_EXCLUIDOS in excel_data:
                    self.df_excluidos = excel_data[NOME_PLANILHA_EXCLUIDOS]; self.df_excluidos = self.df_excluidos.dropna(how='all')
                    for col in COLUNAS_EXCLUIDOS:
                        if col not in self.df_excluidos.columns: self.df_excluidos[col] = pd.NaT if 'Data' in col else pd.NA
                    self.df_excluidos['Validade'] = self._handle_datetime_conversion(self.df_excluidos['Validade'])
                    self.df_excluidos['Data Recebimento'] = self._handle_datetime_conversion(self.df_excluidos['Data Recebimento'])
                    self.df_excluidos['Data Exclusao'] = self._handle_datetime_conversion(self.df_excluidos['Data Exclusao'])
                    self.df_excluidos['Lote'] = self.df_excluidos['Lote'].fillna('').astype(str)
                    self.df_excluidos['Categoria'] = self.df_excluidos['Categoria'].fillna('').astype(str)
                    self.df_excluidos = self.df_excluidos[COLUNAS_EXCLUIDOS]
                else: self.df_excluidos = pd.DataFrame(columns=COLUNAS_EXCLUIDOS)
                self.set_status("Dados carregados.", replace=True)
            except Exception as e:
                messagebox.showerror("Erro ao Ler Excel", f"Não foi possível ler {CAMINHO_PLANILHA}.\nErro: {e}")
                self.df_estoque = pd.DataFrame(columns=COLUNAS_ESTOQUE); self.df_excluidos = pd.DataFrame(columns=COLUNAS_EXCLUIDOS)
                self.set_status("Erro ao carregar dados.", replace=True)
        else:
            self.df_estoque = pd.DataFrame(columns=COLUNAS_ESTOQUE); self.df_excluidos = pd.DataFrame(columns=COLUNAS_EXCLUIDOS)
            self.set_status(f"Arquivo {CAMINHO_PLANILHA} não encontrado.", replace=True)
        # Aplica ordenação inicial
        self._apply_current_sort_to_dataframe()

    def save_data(self) -> None:
        self.set_status(f"Salvando dados em {CAMINHO_PLANILHA}...")
        # ... (código igual à v6) ...
        try:
            with pd.ExcelWriter(CAMINHO_PLANILHA, engine='openpyxl', datetime_format='YYYY-MM-DD', date_format='YYYY-MM-DD') as writer:
                 # Salva Estoque
                df_estoque_save = self.df_estoque.copy()
                for col in ['Validade', 'Data Recebimento']:
                     if col in df_estoque_save.columns:
                          if pd.api.types.is_datetime64_any_dtype(df_estoque_save[col]): df_estoque_save[col] = df_estoque_save[col].dt.strftime('%Y-%m-%d')
                          else: df_estoque_save[col] = pd.to_datetime(df_estoque_save[col], errors='coerce').dt.strftime('%Y-%m-%d')
                     df_estoque_save[col] = df_estoque_save[col].fillna('')
                df_estoque_save['Lote'] = df_estoque_save['Lote'].astype(str); df_estoque_save['Categoria'] = df_estoque_save['Categoria'].astype(str)
                # Salva na ordem atual do DataFrame (que reflete a ordenação da UI)
                df_estoque_save[COLUNAS_ESTOQUE].to_excel(writer, sheet_name=NOME_PLANILHA_ESTOQUE, index=False)

                # Salva Excluídos (ordem pode ser diferente ou padrão)
                df_excluidos_save = self.df_excluidos.copy()
                for col in ['Validade', 'Data Recebimento', 'Data Exclusao']:
                     if col in df_excluidos_save.columns:
                          if pd.api.types.is_datetime64_any_dtype(df_excluidos_save[col]): df_excluidos_save[col] = df_excluidos_save[col].dt.strftime('%Y-%m-%d')
                          else: df_excluidos_save[col] = pd.to_datetime(df_excluidos_save[col], errors='coerce').dt.strftime('%Y-%m-%d')
                     df_excluidos_save[col] = df_excluidos_save[col].fillna('')
                df_excluidos_save['Lote'] = df_excluidos_save['Lote'].astype(str); df_excluidos_save['Categoria'] = df_excluidos_save['Categoria'].astype(str)
                df_excluidos_save[COLUNAS_EXCLUIDOS].to_excel(writer, sheet_name=NOME_PLANILHA_EXCLUIDOS, index=False)

            self.set_status("Dados salvos com sucesso.", replace=True)
        except PermissionError:
             messagebox.showerror("Erro de Permissão ao Salvar", f"Não foi possível salvar {CAMINHO_PLANILHA}."); self.set_status("Falha ao salvar (Permissão negada).", replace=True)
        except Exception as e:
            messagebox.showerror("Erro ao Salvar Excel", f"Não foi possível salvar {CAMINHO_PLANILHA}.\nErro: {e}"); self.set_status("Falha ao salvar.", replace=True)

    # --- Widgets Creation ---
    def create_widgets(self) -> None:
        """Cria todos os elementos da interface gráfica."""
        # ... (Estrutura principal com Notebook igual à v6) ...
        main_frame = ttk.Frame(self.root, padding="5"); main_frame.pack(expand=True, fill=tk.BOTH)
        self.notebook = ttk.Notebook(main_frame); self.notebook.pack(expand=True, fill=tk.BOTH, pady=(0, 5))
        tab_estoque = ttk.Frame(self.notebook, padding="5")
        tab_excluidos = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(tab_estoque, text=" Estoque Ativo "); self.notebook.add(tab_excluidos, text=" Itens Excluídos ")

        # --- Aba Estoque Ativo ---
        # Frame de Entrada (igual v6)
        frame_entrada = ttk.LabelFrame(tab_estoque, text=" Adicionar / Editar Produto ", padding="10")
        frame_entrada.pack(pady=5, fill='x')

        # Linha 0 - Nome
        ttk.Label(frame_entrada, text="Nome:").grid(row=0, column=0, padx=(0,5), pady=2, sticky='w')
        self.entry_nome = ttk.Entry(frame_entrada, width=40)
        self.entry_nome.grid(row=0, column=1, padx=5, pady=2, sticky='ew', columnspan=3)

        # Linha 1 - Categoria
        ttk.Label(frame_entrada, text="Categoria:").grid(row=1, column=0, padx=(0,5), pady=2, sticky='w')
        self.combo_categoria = ttk.Combobox(frame_entrada, values=LISTA_CATEGORIAS, state="readonly", width=20)
        self.combo_categoria.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        # Linha 2 - Validade e Lote
        ttk.Label(frame_entrada, text="Validade:").grid(row=2, column=0, padx=(0,5), pady=2, sticky='w')
        self.entry_validade = ttk.Entry(frame_entrada, width=12)
        self.entry_validade.grid(row=2, column=1, padx=5, pady=2, sticky='w')
        ttk.Label(frame_entrada, text="(DD/MM/AAAA)").grid(row=2, column=2, padx=(0,10), pady=2, sticky='w')
        ttk.Label(frame_entrada, text="Lote:").grid(row=2, column=3, padx=(0,5), pady=2, sticky='w')
        self.entry_lote = ttk.Entry(frame_entrada, width=20)
        self.entry_lote.grid(row=2, column=4, padx=5, pady=2, sticky='w')

        # Linha 3 - Quantidade
        ttk.Label(frame_entrada, text="Quantidade:").grid(row=3, column=0, padx=(0, 5), pady=2, sticky='w')
        self.entry_quantidade = ttk.Entry(frame_entrada, width=12)
        self.entry_quantidade.grid(row=3, column=1, padx=5, pady=2, sticky='w')

        # Linha 4 - Botões
        self.btn_save = ttk.Button(frame_entrada, text="Adicionar Novo Produto", width=25, command=self.save_product)
        self.btn_save.grid(row=4, column=1, padx=5, pady=10, sticky='w')

        self.btn_cancel_edit = ttk.Button(frame_entrada, text="Cancelar Edição", command=self.cancel_edit, state=tk.DISABLED)
        self.btn_cancel_edit.grid(row=4, column=2, padx=5, pady=10, sticky='w', columnspan=2)

        # Ajuste para expandir corretamente
        frame_entrada.columnconfigure(1, weight=1)
        frame_entrada.columnconfigure(4, weight=1)



        # Frame para Ações da Lista e Filtro (NOVO)
        frame_list_controls = ttk.Frame(tab_estoque)
        frame_list_controls.pack(pady=(5, 2), fill='x')

        # Ações (Editar, Mover) - movidos para este frame
        self.btn_edit = ttk.Button(frame_list_controls, text="Editar Selecionado", command=self.load_item_for_edit)
        self.btn_edit.pack(side=tk.LEFT, padx=(0, 5))
        self.btn_remover = ttk.Button(frame_list_controls, text="Mover p/ Excluídos", command=self.move_to_deleted)
        self.btn_remover.pack(side=tk.LEFT, padx=5)

        # Filtro (NOVO) - Adicionado à direita no mesmo frame
        self.btn_clear_search = ttk.Button(frame_list_controls, text="Limpar", width=8, command=self.clear_search)
        self.btn_clear_search.pack(side=tk.RIGHT, padx=(5, 0))
        self.search_entry = ttk.Entry(frame_list_controls, textvariable=self.search_var, width=25)
        self.search_entry.pack(side=tk.RIGHT, padx=5)
        ttk.Label(frame_list_controls, text="Buscar Produto:").pack(side=tk.RIGHT)
        # Evento para filtrar ao digitar
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh_active_stock_view())

        # Lista (Treeview) Estoque Ativo (igual v6, incluindo bind de duplo clique e config de colunas/sort)
        frame_lista_estoque = ttk.Frame(tab_estoque)
        frame_lista_estoque.pack(expand=True, fill='both')
        colunas_estoque_view = ('produto', 'categoria', 'validade', 'lote', 'data_recebimento', 'status')
        self.tree_main = ttk.Treeview(frame_lista_estoque, columns=colunas_estoque_view, show='headings', height=10)
        self.tree_main.bind("<Double-1>", self.on_double_click_edit)
        for col_id in colunas_estoque_view:
            col_text = col_id.replace('_', ' ').title() # Texto padrão
            if col_id == 'data_recebimento': col_text = 'Recebido Em'

         # Cria um dicionário base de argumentos
            heading_args = {
                'text': col_text,
                'anchor': 'center'
            }

            # Adiciona o argumento 'command' SOMENTE se a coluna for ordenável
            if col_id in TREEVIEW_COL_MAP:
                # Cria a função lambda corretamente
                cmd = lambda c=col_id: self.sort_treeview(c)
                heading_args['command'] = cmd
            # Se não for ordenável (ex: 'status'), 'command' não é adicionado

            # Chama heading desempacotando os argumentos do dicionário
            self.tree_main.heading(col_id, **heading_args)
        self.tree_main.column('produto', width=250, anchor='w'); self.tree_main.column('categoria', width=100, anchor='w')
        self.tree_main.column('validade', width=100, anchor='center'); self.tree_main.column('lote', width=110, anchor='center')
        self.tree_main.column('data_recebimento', width=110, anchor='center'); self.tree_main.column('status', width=140, anchor='center')
        scrollbar_main = ttk.Scrollbar(frame_lista_estoque, orient=tk.VERTICAL, command=self.tree_main.yview); self.tree_main.configure(yscroll=scrollbar_main.set)
        self.tree_main.grid(row=0, column=0, sticky='nsew'); scrollbar_main.grid(row=0, column=1, sticky='ns')
        frame_lista_estoque.rowconfigure(0, weight=1); frame_lista_estoque.columnconfigure(0, weight=1)
        self.configure_treeview_tags(self.tree_main)

        # --- Aba Itens Excluídos (igual v6, com botão Restaurar) ---
        frame_botoes_excluidos = ttk.Frame(tab_excluidos); frame_botoes_excluidos.pack(pady=5, fill='x')
        self.btn_restore = ttk.Button(frame_botoes_excluidos, text="Restaurar Selecionado(s)", command=self.restore_deleted_item); self.btn_restore.pack(side=tk.LEFT, padx=(0, 5))
        self.btn_purge = ttk.Button(frame_botoes_excluidos, text="Excluir Permanentemente", command=self.purge_deleted_product); self.btn_purge.pack(side=tk.LEFT, padx=5)
        frame_lista_excluidos = ttk.Frame(tab_excluidos); frame_lista_excluidos.pack(expand=True, fill='both')
        colunas_excluidos_view = ('produto', 'categoria', 'validade', 'lote', 'data_recebimento', 'status', 'data_exclusao')
        self.tree_deleted = ttk.Treeview(frame_lista_excluidos, columns=colunas_excluidos_view, show='headings', height=10)
        for col_id in colunas_excluidos_view:
             col_text = col_id.replace('_', ' ').title();
             if col_id == 'data_recebimento': col_text = 'Recebido Em'
             if col_id == 'data_exclusao': col_text = 'Excluído Em'
             if col_id == 'status': col_text = 'Status na Exclusão'
             self.tree_deleted.heading(col_id, text=col_text, anchor='center')
        self.tree_deleted.column('produto', width=200, anchor='w'); self.tree_deleted.column('categoria', width=90, anchor='w')
        self.tree_deleted.column('validade', width=90, anchor='center'); self.tree_deleted.column('lote', width=100, anchor='center')
        self.tree_deleted.column('data_recebimento', width=100, anchor='center'); self.tree_deleted.column('status', width=130, anchor='center')
        self.tree_deleted.column('data_exclusao', width=100, anchor='center')
        scrollbar_deleted = ttk.Scrollbar(frame_lista_excluidos, orient=tk.VERTICAL, command=self.tree_deleted.yview); self.tree_deleted.configure(yscroll=scrollbar_deleted.set)
        self.tree_deleted.grid(row=0, column=0, sticky='nsew'); scrollbar_deleted.grid(row=0, column=1, sticky='ns')
        frame_lista_excluidos.rowconfigure(0, weight=1); frame_lista_excluidos.columnconfigure(0, weight=1)
        self.configure_treeview_tags(self.tree_deleted)

        # --- Barra de Status (igual v6) ---
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w', padding="5 2")
        status_bar.pack(side=tk.BOTTOM, fill='x')

    # --- Métodos de Controle e Ação ---
    def set_status(self, message: str, replace: bool = True, append: bool = False) -> None:
        """Atualiza a mensagem na barra de status."""
        # ... (código igual à v6) ...
        if hasattr(self, 'status_var'):
            current = self.status_var.get() if append else ""
            prefix = f"{current} | " if append and current else ""
            self.status_var.set(f"{prefix}{message}")
            self.root.update_idletasks()

    def clear_input_fields(self) -> None:
        """Limpa campos de entrada e reseta estado de edição."""
        # ... (código igual à v6) ...
        self.entry_nome.delete(0, tk.END); self.combo_categoria.set('')
        self.entry_validade.delete(0, tk.END); self.entry_lote.delete(0, tk.END)
        if self.editing_index is not None:
            self.editing_index = None; self.btn_save.config(text="Adicionar Novo Produto")
            self.btn_cancel_edit.config(state=tk.DISABLED)
        self.entry_nome.focus()

    def cancel_edit(self) -> None:
        """Cancela o modo de edição."""
        # ... (código igual à v6) ...
        self.editing_index = None; self.clear_input_fields()
        self.set_status("Edição cancelada.")

    def load_item_for_edit(self) -> None:
        """Carrega item selecionado para edição."""
        # ... (código igual à v6) ...
        selected_ids = self.tree_main.selection()
        if not selected_ids: messagebox.showwarning("Nenhum Item", "Selecione um item para editar."); return
        if len(selected_ids) > 1: messagebox.showwarning("Muitos Itens", "Selecione apenas um item."); return
        item_id = selected_ids[0]
        try:
            item_index = int(self.tree_main.item(item_id, 'tags')[0])
            if item_index not in self.df_estoque.index: messagebox.showerror("Erro", "Item não encontrado."); self.refresh_active_stock_view(); return
            item_data = self.df_estoque.loc[item_index]
            self.clear_input_fields(); self.entry_nome.insert(0, item_data['Produto']); self.combo_categoria.set(item_data['Categoria'])
            validade_dt = pd.to_datetime(item_data['Validade']);
            if not pd.isna(validade_dt): self.entry_validade.insert(0, validade_dt.strftime('%d/%m/%Y'))
            self.entry_lote.insert(0, str(item_data.get('Lote', '')))
            self.editing_index = item_index; self.btn_save.config(text="Salvar Alterações")
            self.btn_cancel_edit.config(state=tk.NORMAL); self.set_status(f"Editando: {item_data['Produto']}"); self.entry_nome.focus()
        except (ValueError, IndexError, KeyError, TypeError) as e: messagebox.showerror("Erro ao Carregar", f"Não foi possível carregar.\nErro: {e}"); self.cancel_edit()

    def on_double_click_edit(self, event: tk.Event) -> None:
        """Callback de duplo clique na lista principal."""
        # ... (código igual à v6) ...
        region = self.tree_main.identify_region(event.x, event.y)
        if region == "cell": self.load_item_for_edit()

    def save_product(self) -> None:
        """Salva novo produto ou alterações, com verificação de duplicado."""
        # ... (lógica de obtenção e validação de campos igual v6) ...
        nome = self.entry_nome.get().strip(); categoria = self.combo_categoria.get()
        validade_str = self.entry_validade.get().strip(); lote = self.entry_lote.get().strip()
        if not nome: messagebox.showwarning("Inválido", "Nome."); return
        if not categoria: messagebox.showwarning("Inválido", "Categoria."); return
        if not validade_str: messagebox.showwarning("Inválido", "Validade."); return
        try: validade_dt = datetime.strptime(validade_str, '%d/%m/%Y').replace(hour=23, minute=59, second=59)
        except ValueError: messagebox.showerror("Formato Inválido", "Data: DD/MM/AAAA."); return

        if self.editing_index is not None:
            # Edição
            try:
                nome_original = self.df_estoque.loc[self.editing_index, 'Produto']
                if nome.lower() != nome_original.lower(): # Só checa duplicado se nome mudou
                     mask = (self.df_estoque['Produto'].str.lower() == nome.lower()) & (self.df_estoque.index != self.editing_index)
                     if mask.any():
                          if not messagebox.askyesno("Produto Duplicado", f"Já existe: '{nome}'. Salvar mesmo assim?"): return
                self.df_estoque.loc[self.editing_index, ['Produto', 'Categoria', 'Validade', 'Lote']] = [nome, categoria, validade_dt, lote]
                idx_edited = self.editing_index
                self.save_data(); self.clear_input_fields(); self.refresh_active_stock_view()
                self.set_status(f"Produto índice {idx_edited} atualizado.")
            except KeyError: messagebox.showerror("Erro", "Índice não encontrado."); self.cancel_edit()
            except Exception as e: messagebox.showerror("Erro ao Atualizar", f"{e}"); self.cancel_edit()
        else:
            # Adição
            mask = self.df_estoque['Produto'].str.lower() == nome.lower()
            if mask.any():
                 if not messagebox.askyesno("Produto Duplicado", f"Já existe: '{nome}'. Adicionar mesmo assim?"): return
            data_recebimento = datetime.now().date()
            nova_linha = pd.DataFrame({'Produto': [nome], 'Categoria': [categoria], 'Validade': [validade_dt], 'Lote': [lote], 'Data Recebimento': [data_recebimento]})
            self.df_estoque = pd.concat([self.df_estoque, nova_linha], ignore_index=True)
            self.df_estoque = self.df_estoque.astype({'Validade': 'datetime64[ns]', 'Data Recebimento': 'datetime64[ns]', 'Lote': 'object', 'Categoria': 'object'})
            self._apply_current_sort_to_dataframe() # Reordena após adicionar
            self.save_data(); self.clear_input_fields(); self.refresh_active_stock_view()
            self.set_status(f"Produto '{nome}' adicionado.")

    def move_to_deleted(self) -> None:
        """Move itens selecionados para excluídos."""
        # ... (lógica igual à v6) ...
        selected_ids = self.tree_main.selection()
        if not selected_ids: messagebox.showwarning("Nenhum Item", "Selecione itens."); return
        if not messagebox.askyesno("Confirmar", f"Mover {len(selected_ids)} item(ns)?"): return
        indices_para_mover = [];
        # Dentro do método move_to_deleted:
        indices_para_mover = []
        # (Se você usa rows_para_mover, inicialize aqui também: rows_para_mover = [])

        for item_id in selected_ids:
            try:
                # Pega o índice original do DataFrame da tag do item na Treeview
                idx = int(self.tree_main.item(item_id, 'tags')[0])
                # Verifica se o índice ainda existe no DataFrame atual
                if idx in self.df_estoque.index:
                    indices_para_mover.append(idx)
                    # Se você precisa das linhas (como na v6/v7), adicione aqui:
                    # rows_para_mover.append(self.df_estoque.loc[idx])
                # else: # Opcional: Log se índice não for encontrado
                #    print(f"Aviso: Índice {idx} do item {item_id} não encontrado no df_estoque.")

            except (ValueError, IndexError, KeyError):
                # Ocorre se a tag estiver mal formatada ou vazia
                print(f"Aviso: Não foi possível obter um índice válido para o item {item_id}. Pulando.")
                continue # Pula para o próximo item_id no loop
        rows_para_mover_df = self.df_estoque.loc[indices_para_mover].copy()
        if rows_para_mover_df.empty: self.set_status("Nenhum válido."); return
        rows_para_mover_df['Data Exclusao'] = datetime.now()
        rows_para_mover_df = rows_para_mover_df.reindex(columns=COLUNAS_EXCLUIDOS, fill_value=pd.NaT)
        self.df_excluidos = pd.concat([self.df_excluidos, rows_para_mover_df], ignore_index=True)
        self.df_excluidos = self.df_excluidos.astype({'Validade': 'datetime64[ns]', 'Data Recebimento': 'datetime64[ns]', 'Data Exclusao': 'datetime64[ns]'})
        self.df_excluidos[['Lote', 'Categoria']] = self.df_excluidos[['Lote', 'Categoria']].astype(str)
        self.df_estoque.drop(indices_para_mover, inplace=True); self.df_estoque.reset_index(drop=True, inplace=True)
        if self.editing_index in indices_para_mover: self.cancel_edit()
        self.save_data(); self.refresh_active_stock_view(); self.update_deleted_display()
        self.set_status(f"{len(indices_para_mover)} item(ns) movido(s).")

    def restore_deleted_item(self) -> None:
        """Restaura itens da lixeira para o estoque ativo."""
        # ... (lógica igual à v6) ...
        selected_ids = self.tree_deleted.selection()
        if not selected_ids: messagebox.showwarning("Nenhum Item", "Selecione itens."); return
        if not messagebox.askyesno("Confirmar", f"Restaurar {len(selected_ids)} item(ns)?"): return
        indices_para_restaurar = []
        for item_id in selected_ids: 
            try: 
                idx = int(self.tree_deleted.item(item_id, 'tags')[0]); indices_para_restaurar.append(idx); 
            
            except: continue
        rows_para_restaurar_df = self.df_excluidos.loc[indices_para_restaurar].copy()
        if rows_para_restaurar_df.empty: self.set_status("Nenhum válido."); return
        rows_para_restaurar_df = rows_para_restaurar_df.drop(columns=['Data Exclusao']); rows_para_restaurar_df = rows_para_restaurar_df.reindex(columns=COLUNAS_ESTOQUE, fill_value=pd.NaT)
        self.df_estoque = pd.concat([self.df_estoque, rows_para_restaurar_df], ignore_index=True)
        self.df_estoque = self.df_estoque.astype({'Validade': 'datetime64[ns]', 'Data Recebimento': 'datetime64[ns]'})
        self.df_estoque[['Lote', 'Categoria']] = self.df_estoque[['Lote', 'Categoria']].astype(str)
        self._apply_current_sort_to_dataframe() # Reordena estoque após restaurar
        self.df_excluidos.drop(indices_para_restaurar, inplace=True); self.df_excluidos.reset_index(drop=True, inplace=True)
        self.save_data(); self.refresh_active_stock_view(); self.update_deleted_display()
        self.set_status(f"{len(indices_para_restaurar)} item(ns) restaurado(s).")

    def purge_deleted_product(self) -> None:
        """Exclui permanentemente itens da lixeira."""
        # ... (lógica igual à v6) ...
        selected_ids = self.tree_deleted.selection()
        if not selected_ids: messagebox.showwarning("Nenhum Item", "Selecione itens."); return
        if not messagebox.askyesno("Confirmar Exclusão Permanente", f"Excluir permanentemente {len(selected_ids)} item(ns)?"): return
        indices_para_remover = [];
        for item_id in selected_ids: 
            try: 
                idx = int(self.tree_deleted.item(item_id, 'tags')[0]); indices_para_remover.append(idx); 
            except: 
                continue
        if not indices_para_remover: self.set_status("Nenhum válido."); return
        self.df_excluidos.drop(indices_para_remover, inplace=True); self.df_excluidos.reset_index(drop=True, inplace=True)
        self.save_data(); self.update_deleted_display()
        self.set_status(f"{len(indices_para_remover)} item(ns) excluído(s).")

    # --- Métodos de Display e Formatação ---
    def calculate_status_validade(self, validade: Any) -> Tuple[str, str]:
        """Calcula status e tag de cor."""
        # ... (código igual à v6) ...
        if pd.isna(validade): return "Data Inválida", "vencido"
        hoje = datetime.now().date(); validade_date = pd.to_datetime(validade).date()
        diff_days = (validade_date - hoje).days
        if diff_days < 0: return "Vencido", "vencido"
        elif diff_days <= DIAS_ALERTA_AMARELO: return f"Vence em {diff_days} dia(s)" if diff_days != 0 else "Vence Hoje!", "alerta"
        else: return "OK", "ok"

    def configure_treeview_tags(self, treeview: ttk.Treeview) -> None:
        """Configura cores das tags."""
        # ... (código igual à v6) ...
        treeview.tag_configure('ok', background='lightgreen', foreground='black')
        treeview.tag_configure('alerta', background='#FFFF99', foreground='black')
        treeview.tag_configure('vencido', background='red', foreground='white')

    def _format_date_for_display(self, date_value: Any) -> str:
         """Formata data para DD/MM/YYYY ou retorna ''."""
         # ... (código igual à v6) ...
         if pd.isna(date_value): return ""
         try: return pd.to_datetime(date_value).strftime('%d/%m/%Y')
         except (ValueError, TypeError): return "Erro Data"

    def _apply_current_sort_to_dataframe(self) -> None:
        """Ordena o DataFrame principal (self.df_estoque) in-place."""
        if self.last_sort_col and self.last_sort_col in TREEVIEW_COL_MAP:
             df_col = TREEVIEW_COL_MAP[self.last_sort_col]
             if df_col in self.df_estoque.columns:
                  try:
                       ascending_order = not self.last_sort_reverse
                       # Tenta converter colunas relevantes antes de ordenar
                       if df_col in ['Validade', 'Data Recebimento']:
                            self.df_estoque[df_col] = pd.to_datetime(self.df_estoque[df_col], errors='coerce')
                       elif df_col == 'Lote': # Exemplo: tentar ordenar Lote numericamente se possível
                            # Cuidado: isso pode falhar se Lote tiver letras
                            # self.df_estoque[df_col] = pd.to_numeric(self.df_estoque[df_col], errors='ignore')
                            pass # Manter como string por segurança

                       self.df_estoque = self.df_estoque.sort_values(
                           by=df_col,
                           ascending=ascending_order,
                           na_position='last', # Nulos/NaT no final
                           key=lambda col: col.str.lower() if col.dtype == 'object' else col # Ordenar texto sem case sensitive
                       )
                       self.df_estoque.reset_index(drop=True, inplace=True)
                  except Exception as e: print(f"Erro ao ordenar DF por {df_col}: {e}")

    def _get_filtered_and_sorted_df(self) -> pd.DataFrame:
         """Retorna o DataFrame de estoque filtrado e ordenado para exibição."""
         search_term = self.search_var.get().lower().strip()
         df_to_display = self.df_estoque # Começa com o DF completo

         # Aplica filtro se houver termo de busca
         if search_term:
              try:
                   mask = df_to_display['Produto'].str.lower().str.contains(search_term, na=False)
                   df_to_display = df_to_display[mask]
              except Exception as e:
                   print(f"Erro ao filtrar: {e}") # Log de erro
                   # Continua com o DF não filtrado se der erro

         # Aplica ordenação atual ao DF (completo ou filtrado)
         # A ordenação do DF principal já foi feita por _apply_current_sort_to_dataframe
         # Se aplicamos filtro, precisamos reordenar o subconjunto filtrado
         # No entanto, é mais simples garantir que o DF principal *sempre* reflete a ordenação
         # e filtrar *depois*. Então, _apply_current_sort_to_dataframe já fez o trabalho.

         return df_to_display


    def refresh_active_stock_view(self) -> None:
        """Obtém os dados filtrados/ordenados e atualiza a Treeview principal."""
        df_display = self._get_filtered_and_sorted_df()
        self._populate_main_treeview(df_display)

    def _populate_main_treeview(self, df_to_display: pd.DataFrame) -> None:
        """Limpa e preenche a Treeview principal com dados de um DataFrame."""
        for i in self.tree_main.get_children(): self.tree_main.delete(i)

        # Itera sobre o DataFrame (potencialmente filtrado)
        # MAS usa o índice original (que está no index do df_to_display) como tag
        for index, row in df_to_display.iterrows():
            validade_dt = pd.to_datetime(row.get('Validade'), errors='coerce')
            lote_val = str(row.get('Lote', ''))
            categoria_val = str(row.get('Categoria', ''))
            receb_str = self._format_date_for_display(row.get('Data Recebimento', pd.NaT))
            validade_str = self._format_date_for_display(validade_dt)
            status, tag_cor = self.calculate_status_validade(validade_dt)

            # Passa o índice REAL do DataFrame original como tag[0]
            self.tree_main.insert('', tk.END,
                        values=(row['Produto'], categoria_val, validade_str, lote_val, receb_str, status),
                        tags=(str(index), tag_cor))

    def sort_treeview(self, col_id: str) -> None:
        """Callback para ordenar e atualizar a view principal."""
        if col_id not in TREEVIEW_COL_MAP: self.set_status(f"Não ordena por {col_id.title()}"); return

        reverse = False
        if col_id == self.last_sort_col: reverse = not self.last_sort_reverse
        self.last_sort_col = col_id; self.last_sort_reverse = reverse

        self._apply_current_sort_to_dataframe() # Reordena o DF principal
        self.refresh_active_stock_view() # Atualiza a view (que usará o DF ordenado e o filtro atual)

        direction = "DESC" if self.last_sort_reverse else "ASC"
        self.set_status(f"Ordenado por {TREEVIEW_COL_MAP[col_id]} ({direction})")

    def clear_search(self) -> None:
         """Limpa o campo de busca e atualiza a lista."""
         self.search_var.set("")
         self.refresh_active_stock_view()
         self.set_status("Filtro limpo.")

    # Renomeada para clareza
    # def update_main_display(self): foi substituida por refresh_active_stock_view

    def update_deleted_display(self) -> None:
        """Atualiza a Treeview dos itens excluídos."""
        # ... (lógica igual à v6, talvez reordenar por exclusão aqui) ...
        for i in self.tree_deleted.get_children(): self.tree_deleted.delete(i)
        try:
            self.df_excluidos['Data Exclusao'] = pd.to_datetime(self.df_excluidos['Data Exclusao'])
            df_ordenado = self.df_excluidos.sort_values(by='Data Exclusao', ascending=False)
            df_ordenado.reset_index(drop=True, inplace=True)
        except KeyError: df_ordenado = self.df_excluidos
        for index, row in df_ordenado.iterrows():
            validade_dt = pd.to_datetime(row.get('Validade'), errors='coerce')
            lote_val = str(row.get('Lote', '')); categoria_val = str(row.get('Categoria', ''))
            receb_str = self._format_date_for_display(row.get('Data Recebimento', pd.NaT))
            exclusao_str = self._format_date_for_display(row.get('Data Exclusao', pd.NaT))
            validade_str = self._format_date_for_display(validade_dt)
            status, tag_cor = self.calculate_status_validade(validade_dt)
            self.tree_deleted.insert('', tk.END, values=(row['Produto'], categoria_val, validade_str, lote_val, receb_str, status, exclusao_str), tags=(str(index), tag_cor))

    def run(self) -> None:
        """Inicia o loop principal da aplicação."""
        self.root.mainloop()

# --- Ponto de Entrada ---
if __name__ == "__main__":
    root = tk.Tk()
    app = StockManagerApp(root)
    app.run()


    def aplicar_filtros(self):
        from tkinter import messagebox
        df_filtrado = self.df_estoque.copy()

        categoria = self.categoria_var.get().strip().lower()
        if categoria:
            df_filtrado = df_filtrado[df_filtrado['Categoria'].str.lower().str.contains(categoria)]

        lote = self.lote_var.get().strip().lower()
        if lote:
            df_filtrado = df_filtrado[df_filtrado['Lote'].str.lower().str.contains(lote)]

        validade = self.validade_var.get().strip()
        if validade:
            try:
                data_limite = datetime.strptime(validade, "%d/%m/%Y")
                df_filtrado = df_filtrado[df_filtrado['Validade'] <= data_limite]
            except ValueError:
                messagebox.showerror("Data inválida", "Insira uma data no formato dd/mm/aaaa")

        self.mostrar_estoque(df_filtrado)


def iniciar_aplicacao():
    root = tk.Tk()
    app = StockManagerApp(root)
    app.run()
