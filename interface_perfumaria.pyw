# ==============================================================================
#  INTERFACE GRÁFICA — AutoPreço Perfumaria
# ==============================================================================
#
#  Este arquivo é o PONTO DE ENTRADA do programa (extensão .pyw no Windows
#  faz o Python rodar SEM abrir a janela do terminal — só a GUI aparece).
#
#  Responsabilidades:
#    - Janela principal com campos de entrada (CSV, XLSX, cotação, frete)
#    - Log visual colorido de processamento
#    - Preview modal antes de salvar (com troca de match e remoção de linhas)
#    - Reassociação manual de produtos não encontrados
#    - Pergunta individual para variações de preço pequenas
#
#  ARQUITETURA:
#    Tkinter é single-threaded — se rodarmos o processamento na thread
#    principal, a interface TRAVA até terminar. Por isso usamos:
#      - threading.Thread para o processamento pesado
#      - queue.Queue para comunicação thread-safe entre processamento e UI
#      - root.after() para drenar a fila de mensagens na main thread
#
#  Nota para o júnior: NUNCA atualize widgets Tkinter de uma thread secundária!
#  Isso causa crashes aleatórios. Sempre use uma fila + after() para delegar
#  a atualização para a main thread.
# ==============================================================================

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import threading
import queue
from logica_perfumaria import (processar_perfumaria, salvar_resultado,
                                aplicar_reassociacao, arredondar_preco,
                                formatar_moeda, parse_moeda,
                                gerar_descricao_fornecedor)

# ──────────────────────────────────────────────────────────────────────────────
#  JANELA DE PREVIEW  (com menu de contexto para trocar match)
# ──────────────────────────────────────────────────────────────────────────────
#
#  Esta janela é aberta APÓS o processamento, ANTES de salvar o CSV.
#  Mostra uma tabela com todos os produtos que serão exportados, permitindo
#  ao usuário:
#    - Conferir visualmente cada match (XLSX ↔ CSV)
#    - Trocar um match errado (duplo clique ou clique direito)
#    - Remover uma linha da exportação (clique direito → Remover)
#    - Confirmar e exportar ou cancelar
#
#  O código de cores indica a sensibilidade da alteração:
#    - Branco   → preço não mudou
#    - Amarelo  → variação ≤ R$10 (pequena, provavelmente ok)
#    - Laranja  → variação R$10–R$80 (atenção)
#    - Vermelho → variação ≥ R$80 (perigo, revisar!)
#    - Azul     → match corrigido manualmente (reassociado)
# ──────────────────────────────────────────────────────────────────────────────

class PreviewWindow(tk.Toplevel):
    # Cores da paleta escura (dark mode)
    COR_BG       = "#12122a"      # Fundo da janela
    COR_HEADER   = "#1e1e3f"      # Fundo do cabeçalho da tabela
    COR_NORMAL   = "#ccccdd"      # Texto padrão (cinza claro)
    COR_GTIN_AMAR= "#ffd700"      # Amarelo — variação ≤ R$10
    COR_GTIN_LAR = "#ff8c00"      # Laranja — variação R$10–R$80
    COR_GTIN_VERM= "#ff3b3b"      # Vermelho — variação ≥ R$80
    COR_REASSOC  = "#66aaff"      # Azul — linha alterada manualmente

    def __init__(self, parent, resultado, log_callback):
        """
        Inicializa a janela de Preview.

        Parâmetros:
          parent       — janela pai (self.root da PerfumariaApp)
          resultado    — dicionário retornado por processar_perfumaria()
          log_callback — função log() da PerfumariaApp para escrever no log
        """
        super().__init__(parent)
        self.resultado    = resultado
        self.log_callback = log_callback
        self.confirmado   = False     # Será True se o usuário clicar "Confirmar"

        modo = resultado['modo']
        self.modo = modo
        self.title(f"👁  Preview — {'Modo Estoque' if modo == 'estoque' else 'Atualização de Preços'}")
        self.geometry("980x560")
        self.configure(bg=self.COR_BG)

        # grab_set() torna esta janela MODAL — o usuário não pode interagir
        # com a janela principal enquanto o preview estiver aberto.
        self.grab_set()
        # lift() e focus_force() garantem que a janela fique ACIMA de todas
        # as outras e receba o foco. Sem isso, no Windows a janela pode
        # abrir atrás da principal (bug relatado na Fase 8).
        self.lift()
        self.focus_force()
        self.resizable(True, True)

        self._criar_widgets(modo)
        self._popular_tabela(modo)
        self._verificar_pode_confirmar()

    # ── Mapeamento item_tree → preview_row (para edição) ──────────────────────
    def _item_para_row_idx(self, item_id):
        """
        Converte o ID interno do Treeview para o índice em preview_rows.

        O Treeview do Tkinter usa IDs internos (strings como "I001", "I002")
        que não correspondem diretamente ao índice da lista. Precisamos
        descobrir a posição do item na lista de filhos para acessar
        o dicionário correto em preview_rows.
        """
        todos = self.tree.get_children()
        return list(todos).index(item_id)

    # ── Cor da linha baseada em GTIN + diff ───────────────────────────────────
    def _tag_para_row(self, row):
        """
        Determina a tag de cor para uma linha do preview.

        A lógica segue a mesma escala definida na logica_perfumaria.py:
          - Reassociado manualmente → azul (tem prioridade sobre tudo)
          - Sem GTIN ou diff == 0   → normal (branco/cinza)
          - diff ≤ 10               → amarelo
          - diff < 80               → laranja
          - diff ≥ 80               → vermelho
        """
        if row.get('reassociado'):
            return "reassociado"
        gtin       = row.get('gtin', '') or ''
        diff_preco = row.get('diff_preco', 0) or 0
        if not gtin or diff_preco == 0:
            return "normal"
        if diff_preco <= 10:
            return "gtin_amarelo"
        if diff_preco < 80:
            return "gtin_laranja"
        return "gtin_vermelho"

    def _criar_widgets(self, modo):
        """
        Cria todos os widgets da janela de preview:
          - Header com contagem de produtos e alerta de GTIN
          - Barra de desconto global (somente modo preço)
          - Treeview (tabela) com scrollbars
          - Botões "Confirmar" e "Cancelar"
          - Legenda de cores

        Funcionalidades de edição de preço:
          - Desconto global: campo % na barra superior → recalcula todos
            via arredondar_preco(custo×2 × (1 - pct/100))
          - Edição individual: duplo clique na coluna "Preço Novo" OU
            clique direito → "Editar preço"
        """
        # ── Cabeçalho ─────────────────────────────────────────────────
        header = tk.Frame(self, bg="#0d0d1a", pady=10)
        header.pack(fill=tk.X)

        rows   = self.resultado['preview_rows']
        n_gtin = sum(1 for r in rows if r.get('gtin'))

        tk.Label(header,
                 text=f"👁  Preview de Exportação — {len(rows)} produto(s)",
                 font=("Helvetica", 13, "bold"),
                 bg="#0d0d1a", fg="#c9a8ff").pack()
        if n_gtin:
            tk.Label(header,
                     text=f"⚠️  {n_gtin} produto(s) com GTIN/EAN — troca de preço afeta comparadores",
                     font=("Segoe UI", 9),
                     bg="#0d0d1a", fg=self.COR_GTIN_LAR).pack()
        tk.Label(header,
                 text="Duplo clique na coluna 'Preço Novo' para editar.  Clique direito para mais opções.",
                 font=("Segoe UI", 8), bg="#0d0d1a", fg="#666688").pack()

        # ── Barra de desconto global (somente modo preço) ─────────────
        #  Permite aplicar um desconto percentual em TODOS os produtos
        #  de uma vez. O cálculo respeita a fórmula original:
        #    preço = arredondar_preco(custo × 2 × (1 - pct/100))
        #  Ficará oculta no modo estoque.
        if modo != 'estoque':
            frame_desc = tk.Frame(self, bg="#1a1a35", pady=6)
            frame_desc.pack(fill=tk.X, padx=12, pady=(0, 4))

            tk.Label(frame_desc, text="🏷️  Desconto global:",
                     bg="#1a1a35", fg="#c9a8ff",
                     font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(6, 4))

            self._var_desconto = tk.StringVar(value="0")
            entry_desc = tk.Entry(frame_desc, textvariable=self._var_desconto,
                                  width=6, font=("Consolas", 10),
                                  bg="#0d0d1a", fg="#ffffff",
                                  insertbackground="white",
                                  relief="flat", bd=2,
                                  justify="center")
            entry_desc.pack(side=tk.LEFT, padx=(0, 2))
            tk.Label(frame_desc, text="%",
                     bg="#1a1a35", fg="#aaaacc",
                     font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 10))

            tk.Button(frame_desc, text="Aplicar a todos",
                      bg="#7c3aed", fg="white",
                      font=("Segoe UI", 9, "bold"),
                      relief="flat", cursor="hand2", padx=10, pady=3,
                      activebackground="#6d28d9",
                      command=self._aplicar_desconto).pack(side=tk.LEFT)

            tk.Label(frame_desc,
                     text="  ex: 20 → aplica 20% de desconto antes do arredondamento para ,90",
                     bg="#1a1a35", fg="#555577",
                     font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=8)
        else:
            self._var_desconto = None

        # ── Treeview (tabela de dados) ────────────────────────────────
        #  O Treeview do ttk é o widget "tabela" do Tkinter. Ele suporta
        #  colunas, ordenação e scrollbars. Aqui usamos o estilo "Dark"
        #  customizado para combinar com a paleta escura da interface.
        frame_tree = tk.Frame(self, bg=self.COR_BG, padx=12, pady=6)
        frame_tree.pack(fill=tk.BOTH, expand=True)

        # Colunas diferentes para cada modo — "sel" é sempre a primeira (checkbox)
        if modo == 'estoque':
            colunas    = ("sel", "sku", "descricao", "antes", "depois", "gtin")
            cabecalhos = {
                "sel"      : ("✔",             30),
                "sku"      : ("SKU",           70),
                "descricao": ("Descrição CSV", 340),
                "antes"    : ("Est. Atual",    100),
                "depois"   : ("Est. Novo",     100),
                "gtin"     : ("GTIN/EAN",      140),
            }
        else:
            colunas    = ("sel", "sku", "nome_xlsx", "descricao", "antes", "depois", "custo", "gtin")
            cabecalhos = {
                "sel"      : ("✔",             30),
                "sku"      : ("SKU",           65),
                "nome_xlsx": ("Nome XLSX",     200),
                "descricao": ("Descrição CSV", 200),
                "antes"    : ("Preço Atual",    95),
                "depois"   : ("Preço Novo",     95),
                "custo"    : ("Custo",           85),
                "gtin"     : ("GTIN/EAN",       120),
            }

        # Configuração do estilo escuro para o Treeview.
        # ttk.Style permite customizar a aparência de widgets ttk sem
        # subclassear. Usamos o tema "clam" como base porque é o que
        # melhor aceita customização em todas as plataformas.
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.Treeview",
            background="#0d0d1a", foreground=self.COR_NORMAL,
            fieldbackground="#0d0d1a", rowheight=24,
            font=("Consolas", 9))
        style.configure("Dark.Treeview.Heading",
            background=self.COR_HEADER, foreground="#aaaaff",
            font=("Segoe UI", 9, "bold"))
        style.map("Dark.Treeview", background=[("selected", "#3a3a6e")])

        self.tree = ttk.Treeview(frame_tree, columns=colunas,
                                  show="headings", style="Dark.Treeview")
        for col, (label, width) in cabecalhos.items():
            self.tree.heading(col, text=label)
            anchor = "center" if col == "sel" else "w"
            self.tree.column(col, width=width, anchor=anchor, minwidth=width)

        # Scrollbars vertical e horizontal
        sb_v = ttk.Scrollbar(frame_tree, orient="vertical",   command=self.tree.yview)
        sb_h = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h.grid(row=1, column=0, sticky="ew")
        frame_tree.rowconfigure(0, weight=1)
        frame_tree.columnconfigure(0, weight=1)

        # Configura as tags de cor — cada tag define a cor do texto da linha.
        # Quando inserimos um item com tags=("gtin_amarelo",), o texto fica amarelo.
        self.tree.tag_configure("normal",       foreground=self.COR_NORMAL)
        self.tree.tag_configure("gtin_amarelo", foreground=self.COR_GTIN_AMAR)
        self.tree.tag_configure("gtin_laranja", foreground=self.COR_GTIN_LAR)
        self.tree.tag_configure("gtin_vermelho",foreground=self.COR_GTIN_VERM)
        self.tree.tag_configure("reassociado",  foreground=self.COR_REASSOC)
        self.tree.tag_configure("desmarcado",   foreground="#555566")

        # ── Bindings para interação ──────────────────────────────────
        #  <Button-1>  = clique esquerdo — toggle checkbox se na coluna sel
        #  <Double-1>  = duplo clique — trocar match OU editar valor
        #  <Button-3>  = clique direito — menu de contexto
        self.tree.bind("<Button-1>",    self._toggle_selecao)
        self.tree.bind("<Double-1>",    self._on_duplo_clique)
        self.tree.bind("<Button-3>",    self._menu_contexto)

        # Menu de contexto (clique direito) — aparece na posição do cursor
        self._menu = tk.Menu(self, tearoff=0, bg="#1e1e3f", fg="white",
                             activebackground="#3a3a6e")
        self._menu.add_command(label="✏️  Trocar produto do CSV",
                               command=self._abrir_trocar_match)
        if modo != 'estoque':
            self._menu.add_command(label="💲  Editar preço desta linha",
                                   command=self._editar_via_menu)
        self._menu.add_command(label="🗑  Remover desta exportação",
                               command=self._remover_linha)

        # ── Botões de ação ────────────────────────────────────────────
        frame_btn = tk.Frame(self, bg=self.COR_BG, pady=10)
        frame_btn.pack(fill=tk.X, padx=12)

        self.btn_confirmar = tk.Button(
            frame_btn, text="✅  Confirmar e Exportar",
            bg="#15803d", fg="white", font=("Arial", 11, "bold"),
            relief="flat", cursor="hand2", padx=18, pady=6,
            activebackground="#166534",
            command=self._confirmar)
        self.btn_confirmar.pack(side=tk.LEFT, padx=(0, 12))

        tk.Button(frame_btn, text="❌  Cancelar",
                  bg="#7f1d1d", fg="white", font=("Arial", 11),
                  relief="flat", cursor="hand2", padx=18, pady=6,
                  activebackground="#991b1b",
                  command=self.destroy).pack(side=tk.LEFT)

        tk.Label(frame_btn,
                 text="✏️ duplo clique = trocar match / editar valor   🗑 direito = remover   ☑ clique = marcar/desmarcar",
                 bg=self.COR_BG, fg="#555577",
                 font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=16)

        # Label de aviso (conflitos, etc.) — inicialmente vazio
        self.lbl_aviso = tk.Label(self, text="", bg=self.COR_BG, fg="#ff4d4d",
                                  font=("Segoe UI", 9, "bold"))
        self.lbl_aviso.pack(fill=tk.X, padx=12)

        # ── Legenda de cores (só aparece se há produtos com GTIN) ─────
        if any(r.get('gtin') for r in self.resultado['preview_rows']):
            frame_leg = tk.Frame(frame_btn, bg=self.COR_BG)
            frame_leg.pack(side=tk.RIGHT, padx=8)
            tk.Label(frame_leg, text="GTIN/EAN:", bg=self.COR_BG,
                     fg="#aaaacc", font=("Segoe UI", 8, "bold")).pack(anchor="e")
            for cor, txt in [
                (self.COR_NORMAL,    "● Sem alteração"),
                (self.COR_GTIN_AMAR, "● Variação ≤ R$10"),
                (self.COR_GTIN_LAR,  "● Variação R$10–R$80"),
                (self.COR_GTIN_VERM, "● Variação ≥ R$80"),
                (self.COR_REASSOC,   "● Corrigido manualmente"),
            ]:
                tk.Label(frame_leg, text=txt, bg=self.COR_BG,
                         fg=cor, font=("Segoe UI", 8)).pack(anchor="e")

    def _popular_tabela(self, modo):
        """
        Popula o Treeview com os dados de preview_rows.

        Cada linha recebe uma tag de cor baseada no GTIN e na variação de preço.
        Chamado na inicialização e pode ser chamado novamente para refresh.
        """
        self.tree.delete(*self.tree.get_children())
        for row in self.resultado['preview_rows']:
            tag          = self._tag_para_row(row)
            gtin_display = (row.get('gtin') or '—')
            sel_char     = row.get('sel', '☑')
            if modo == 'estoque':
                valores = (sel_char, row['sku'], row['descricao'],
                           row['antes'], row['depois'], gtin_display)
            else:
                nome_xlsx = row.get('nome_xlsx') or row.get('nome_perfume') or '—'
                valores = (sel_char, row['sku'], nome_xlsx, row['descricao'],
                           row['antes'], row['depois'],
                           row.get('custo', ''), gtin_display)
            tags = [tag]
            if sel_char == '☐':
                tags.append('desmarcado')
            self.tree.insert("", tk.END, values=valores, tags=tuple(tags))

    # ── Desconto global (barra de desconto) ───────────────────────────────────
    def _aplicar_desconto(self):
        """
        Aplica um desconto percentual em TODOS os produtos do preview.

        Fórmula:
          preço_novo = arredondar_preco(custo × 2 × (1 - pct/100))

        O campo 'custo' em preview_rows contém o custo calculado pela
        Passada 2 (valor_unitario × cotação × (1 + frete%)).
        Multiplicando por 2 e aplicando o desconto antes do arredondamento
        para ,90, mantemos a consistência com a fórmula original.

        Se algum produto não tiver 'custo' armazenado (ex: reassociado
        sem custo), usa o 'depois' atual como fallback.
        """
        pct_str = (self._var_desconto.get() if self._var_desconto else '0').strip()
        pct_str = pct_str.replace(',', '.')
        try:
            pct = float(pct_str)
        except ValueError:
            messagebox.showwarning(
                "Desconto inválido",
                f"'{pct_str}' não é um número válido.\n"
                f"Digite um número como 20 ou 15,5.",
                parent=self
            )
            return
        if not (0 <= pct < 100):
            messagebox.showwarning(
                "Desconto inválido",
                "O desconto deve estar entre 0% e 99,9%.",
                parent=self
            )
            return

        fator = 1.0 - pct / 100.0
        df_bling = self.resultado['df_bling']
        alterados = 0

        for r in self.resultado['preview_rows']:
            custo_str = r.get('custo', '')
            if custo_str:
                custo_num = parse_moeda(str(custo_str))
            else:
                # Fallback: tenta inferir o custo a partir do preço atual/2
                custo_num = parse_moeda(str(r.get('depois', '0'))) / 2.0

            if custo_num <= 0:
                continue  # Sem custo válido, pula

            preco_base      = custo_num * 2          # markup ×2
            preco_com_desc  = preco_base * fator      # aplica desconto
            preco_final     = arredondar_preco(preco_com_desc)  # arredonda para ,90
            preco_fmt       = formatar_moeda(preco_final)

            preco_antigo_num = parse_moeda(str(r.get('antes', '0')))
            diff             = abs(preco_final - preco_antigo_num)

            r['depois']     = preco_fmt
            r['diff_preco'] = diff

            # Atualiza df_bling para que a exportação use o novo preço
            csv_idx = r.get('csv_idx')
            if csv_idx is not None and 'Preço' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço'] = preco_fmt

            alterados += 1

        self._popular_tabela(self.modo)
        self._verificar_pode_confirmar()

        msg = (f"✂️  Desconto de {pct:.1f}% aplicado a {alterados} produto(s)."
               if pct > 0
               else f"↩️  Desconto removido — preços recalculados sem desconto ({alterados} produto(s)).")
        self.log_callback(msg, "alerta")

    # ── Editar preço via menu de contexto ─────────────────────────────────────
    def _editar_via_menu(self):
        """
        Atalho para _editar_valor_inline usando a linha atualmente selecionada.
        Chamado pelo item 'Editar preço desta linha' no menu de contexto.
        """
        sel = self.tree.selection()
        if sel:
            self._editar_valor_inline(sel[0])

    # ── Menu contexto ─────────────────────────────────────────────────────────
    def _menu_contexto(self, event):
        """
        Exibe o menu de contexto na posição do cursor quando o usuário
        clica com o botão direito em uma linha.

        Nota: selecionamos a linha ANTES de mostrar o menu para que o
        usuário saiba qual linha será afetada pela ação escolhida.
        """
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self._menu.post(event.x_root, event.y_root)

    # ── Remover linha ─────────────────────────────────────────────────────────
    def _remover_linha(self):
        """
        Remove a linha selecionada da exportação.

        Remove de 3 lugares:
          1. preview_rows (lista de dicts para a tabela)
          2. indices_alterados (set de índices do CSV que serão exportados)
          3. Treeview visual (a linha desaparece da tabela)

        O produto NÃO será exportado no CSV final.
        """
        sel = self.tree.selection()
        if not sel:
            return
        item_id  = sel[0]
        row_idx  = self._item_para_row_idx(item_id)
        row      = self.resultado['preview_rows'][row_idx]
        col_sku  = self.resultado.get('col_sku')
        df_bling = self.resultado['df_bling']

        # Remove de preview_rows e de indices_alterados
        self.resultado['preview_rows'].pop(row_idx)
        if col_sku:
            # Filtra indices_alterados removendo o índice cujo SKU bate com o removido.
            # Usamos set comprehension para manter a estrutura de set.
            self.resultado['indices_alterados'] = {
                idx for idx in self.resultado['indices_alterados']
                if str(df_bling.at[idx, col_sku]).strip() != row['sku']
            }
        self.tree.delete(item_id)
        self.log_callback(f"ℹ️  Removido da exportação: SKU {row['sku']} | {row['descricao']}", "alerta")

    # ── Toggle seleção (checkbox) ──────────────────────────────────────────────
    def _toggle_selecao(self, event):
        """
        Alterna entre ☑ e ☐ quando o usuário clica na coluna de seleção.

        Identifica a coluna clicada. Se for a coluna #1 (sel), alterna o
        estado e aplica/remove a tag 'desmarcado'. Chama
        _verificar_pode_confirmar() após cada toggle.
        """
        # Identifica coluna e linha clicadas
        coluna = self.tree.identify_column(event.x)
        item   = self.tree.identify_row(event.y)
        if not item or coluna != '#1':
            return  # Clique fora da coluna sel — não faz nada

        row_idx  = self._item_para_row_idx(item)
        row      = self.resultado['preview_rows'][row_idx]
        valores  = list(self.tree.item(item, 'values'))
        tags_atuais = list(self.tree.item(item, 'tags'))

        if valores[0] == '☑':
            valores[0] = '☐'
            row['sel'] = '☐'
            if 'desmarcado' not in tags_atuais:
                tags_atuais.append('desmarcado')
        else:
            valores[0] = '☑'
            row['sel'] = '☑'
            if 'desmarcado' in tags_atuais:
                tags_atuais.remove('desmarcado')

        self.tree.item(item, values=tuple(valores), tags=tuple(tags_atuais))
        self._verificar_pode_confirmar()

    # ── Duplo clique: trocar match OU editar valor ───────────────────────────
    def _on_duplo_clique(self, event):
        """
        Handler de duplo clique. Decide se abre o diálogo de troca de
        match ou a edição inline de valor, dependendo da coluna clicada.

        Se o clique for na coluna "depois" (Preço Novo / Est. Novo),
        abre edição inline. Caso contrário, abre TrocarMatchDialog.
        """
        coluna = self.tree.identify_column(event.x)
        item   = self.tree.identify_row(event.y)
        if not item:
            return

        # Descobre qual coluna lógica corresponde ao #N clicado
        colunas = self.tree['columns']
        col_num = int(coluna.replace('#', '')) - 1  # 0-indexed
        if 0 <= col_num < len(colunas):
            col_id = colunas[col_num]
        else:
            col_id = ''

        if col_id == 'depois':
            self._editar_valor_inline(item)
        elif col_id != 'sel':  # Não abre trocar match se clicou no checkbox
            self.tree.selection_set(item)
            self._abrir_trocar_match(event)

    # ── Edição inline de valor (preço/estoque) ───────────────────────────────
    def _editar_valor_inline(self, item_id):
        """
        Abre um simpledialog para editar o valor da coluna "depois"
        (Preço Novo ou Est. Novo) de uma linha específica.

        Valida que o input é numérico, atualiza preview_rows e df_bling,
        e re-renderiza a linha no treeview.
        """
        row_idx = self._item_para_row_idx(item_id)
        row     = self.resultado['preview_rows'][row_idx]

        titulo = "Editar Estoque" if self.modo == 'estoque' else "Editar Preço"
        prompt = "Novo estoque:" if self.modo == 'estoque' else "Novo preço (R$):"

        novo_valor = simpledialog.askstring(
            titulo, prompt,
            initialvalue=str(row['depois']),
            parent=self
        )
        if novo_valor is None:  # Cancelou
            return

        # Valida: aceita apenas decimais (ponto ou vírgula)
        novo_valor = novo_valor.strip().replace(',', '.')
        try:
            valor_num = float(novo_valor)
        except ValueError:
            messagebox.showwarning("Valor Inválido",
                                   f"'{novo_valor}' não é um número válido.",
                                   parent=self)
            return

        # Atualiza dados
        if self.modo == 'estoque':
            fmt = lambda v: str(int(v)) if v == int(v) else f"{v:.2f}".replace('.', ',')
            valor_formatado = fmt(valor_num)
            row['depois'] = valor_formatado
            # Atualiza df_bling
            csv_idx = row.get('csv_idx')
            if csv_idx is not None and 'Estoque' in self.resultado['df_bling'].columns:
                self.resultado['df_bling'].at[csv_idx, 'Estoque'] = valor_formatado
        else:
            valor_formatado = formatar_moeda(valor_num)
            preco_antigo = parse_moeda(str(row['antes']))
            diff_preco = abs(valor_num - preco_antigo)
            row['depois'] = valor_formatado
            row['diff_preco'] = diff_preco
            # Atualiza df_bling
            csv_idx = row.get('csv_idx')
            if csv_idx is not None and 'Preço' in self.resultado['df_bling'].columns:
                self.resultado['df_bling'].at[csv_idx, 'Preço'] = valor_formatado

        # Re-renderiza a linha no treeview
        tag = self._tag_para_row(row)
        gtin_display = (row.get('gtin') or '—')
        sel_char = row.get('sel', '☑')
        if self.modo == 'estoque':
            novos_valores = (sel_char, row['sku'], row['descricao'],
                             row['antes'], row['depois'], gtin_display)
        else:
            nome_xlsx_upd = (row.get('nome_xlsx')
                             or row.get('nome_perfume') or '—')
            novos_valores = (sel_char, row['sku'], nome_xlsx_upd,
                             row['descricao'],
                             row['antes'], row['depois'],
                             row.get('custo', ''), gtin_display)
        tags = [tag]
        if sel_char == '☐':
            tags.append('desmarcado')
        self.tree.item(item_id, values=novos_valores, tags=tuple(tags))

        self.log_callback(
            f"✏️  Valor editado: SKU {row['sku']} → {valor_formatado}",
            "info"
        )
        self._verificar_pode_confirmar()

    # ── Trocar match ──────────────────────────────────────────────────────────
    def _abrir_trocar_match(self, event=None):
        """
        Abre o diálogo TrocarMatchDialog para a linha selecionada.

        Após o diálogo fechar, se o match foi confirmado, atualiza os
        valores visuais na Treeview para refletir a troca.
        """
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        row_idx = self._item_para_row_idx(item_id)
        row     = self.resultado['preview_rows'][row_idx]

        # Abre o diálogo e ESPERA ele fechar (wait_window bloqueia)
        dlg = TrocarMatchDialog(self, self.resultado, row, row_idx, self.log_callback)
        self.wait_window(dlg)

        if dlg.confirmado:
            # Atualiza a linha visual na Treeview com os novos valores
            row_atualizado = self.resultado['preview_rows'][row_idx]
            tag = self._tag_para_row(row_atualizado)
            gtin_display = (row_atualizado.get('gtin') or '—')
            sel_char = row_atualizado.get('sel', '☑')
            if self.modo == 'estoque':
                novos_valores = (sel_char, row_atualizado['sku'],
                                 row_atualizado['descricao'],
                                 row_atualizado['antes'], row_atualizado['depois'],
                                 gtin_display)
            else:
                nome_xlsx_upd = (row_atualizado.get('nome_xlsx')
                                 or row_atualizado.get('nome_perfume') or '—')
                novos_valores = (sel_char, row_atualizado['sku'], nome_xlsx_upd,
                                 row_atualizado['descricao'],
                                 row_atualizado['antes'], row_atualizado['depois'],
                                 row_atualizado.get('custo', ''), gtin_display)
            tags = [tag]
            if sel_char == '☐':
                tags.append('desmarcado')
            self.tree.item(item_id, values=novos_valores, tags=tuple(tags))
            self._verificar_pode_confirmar()

    # ── Verificação de conflitos e habilitação do botão ────────────────────────
    def _verificar_pode_confirmar(self):
        """
        Verifica se o botão Confirmar deve ser habilitado ou desabilitado.

        Condições de bloqueio do botão:
          1. Conflito N:1 — dois ou mais marcados (☑) apontam para o mesmo csv_idx
          (Não encontrados bloqueiam a exportação em _confirmar, não aqui, 
           para que o usuário possa clicar e ver a lista no messagebox)

        Atualiza self.btn_confirmar e self.lbl_aviso.
        """
        avisos = []
        bloquear_botao = False

        # ── 1. Conta linhas marcadas e verifica conflitos N:1 ──
        marcados_csv_idx = []  # lista de csv_idx das linhas marcadas
        n_marcados = 0
        for row in self.resultado['preview_rows']:
            if row.get('sel', '☑') == '☑':
                n_marcados += 1
                csv_idx = row.get('csv_idx')
                if csv_idx is not None:
                    marcados_csv_idx.append(csv_idx)

        # Detecta duplicatas (N:1)
        vistos = {}
        conflitos = []
        for idx in marcados_csv_idx:
            vistos[idx] = vistos.get(idx, 0) + 1
        for idx, count in vistos.items():
            if count > 1:
                conflitos.append(idx)

        if conflitos:
            avisos.append(
                f"⚠️ Conflito N:1 — {len(conflitos)} produto(s) do CSV com múltiplas associações marcadas. "
                f"Desmarque uma das duplicatas."
            )
            bloquear_botao = True

        # ── 2. Não encontrados não resolvidos (apenas modo preço) ──
        if self.modo != 'estoque':
            nao_enc = self.resultado.get('nao_encontrados', [])
            if nao_enc:
                avisos.append(
                    f"⚠️ {len(nao_enc)} produto(s) sem associação ao CSV. A exportação será bloqueada (clique Confirmar para ver a lista)."
                )

        # Atualiza UI
        if avisos:
            self.lbl_aviso.config(text="\n".join(avisos))
        else:
            self.lbl_aviso.config(text=f"{n_marcados} produto(s) marcado(s) para exportação." if n_marcados else "")

        if bloquear_botao:
            self.btn_confirmar.config(state=tk.DISABLED, bg="#3a3a4a", fg="#888899")
        else:
            self.btn_confirmar.config(state=tk.NORMAL, bg="#15803d", fg="white")

    def _confirmar(self):
        """
        Chamado quando o usuário clica "✅ Confirmar e Exportar".

        Antes de salvar:
          - Verifica não encontrados (modo preço)
          - Loga e remove linhas desmarcadas (☐)

        Salva o CSV via salvar_resultado() e fecha a janela.
        O atributo `confirmado` é verificado pela PerfumariaApp após
        o wait_window() para saber se o usuário confirmou ou cancelou.
        """
        # ── Item 6: Bloqueio se há não-associados (modo preço) ─────
        if self.modo != 'estoque':
            nao_enc = self.resultado.get('nao_encontrados', [])
            if nao_enc:
                messagebox.showerror(
                    "Exportação Bloqueada",
                    f"{len(nao_enc)} produto(s) não foram associados ao CSV:\n\n"
                    + "\n".join(f"  • {nome}" for nome in nao_enc[:15])
                    + (f"\n  ... e mais {len(nao_enc) - 15}" if len(nao_enc) > 15 else ""),
                    parent=self
                )
                return

        # ── Item 4: Log e remoção dos desmarcados ──────────────────
        col_sku  = self.resultado.get('col_sku')
        df_bling = self.resultado['df_bling']
        linhas_manter = []
        for row in self.resultado['preview_rows']:
            if row.get('sel', '☑') == '☐':
                # Linha desmarcada — loga e remove do indices_alterados
                self.log_callback(
                    f"ℹ️ EXCLUÍDO [SKU {row['sku']}]: {row['descricao']}",
                    "alerta"
                )
                csv_idx = row.get('csv_idx')
                if csv_idx is not None:
                    self.resultado['indices_alterados'].discard(csv_idx)
                elif col_sku:
                    self.resultado['indices_alterados'] = {
                        idx for idx in self.resultado['indices_alterados']
                        if str(df_bling.at[idx, col_sku]).strip() != row['sku']
                    }
            else:
                linhas_manter.append(row)
        self.resultado['preview_rows'] = linhas_manter

        salvar_resultado(self.resultado, self.log_callback)
        self.confirmado = True
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
#  DIÁLOGO DE TROCA DE MATCH  (acessado pelo duplo clique no preview)
# ──────────────────────────────────────────────────────────────────────────────
#
#  Este diálogo permite ao usuário CORRIGIR um match errado do algoritmo.
#
#  Exemplo de uso:
#    O Jaccard associou "Aseel Maison Asrar" ao produto "Rey Maison Asrar"
#    (errado!). O usuário faz duplo clique na linha do preview, e este diálogo
#    abre mostrando:
#      - O nome do XLSX ("Aseel Maison Asrar")
#      - O match atual ("Rey Maison Asrar")
#      - Um combobox filtrável com TODOS os produtos do CSV
#
#    O usuário pode:
#      1. Digitar no combobox para filtrar e escolher manualmente
#      2. Clicar "🔄 Ignorar e tentar novamente" — roda o Jaccard excluindo
#         o candidato atual (e todos os anteriormente rejeitados)
#      3. Repetir quantas vezes quiser (set _excluidos cresce a cada tentativa)
# ──────────────────────────────────────────────────────────────────────────────

class TrocarMatchDialog(tk.Toplevel):
    COR_BG     = "#12122a"
    COR_HEADER = "#1e1e3f"

    def __init__(self, parent, resultado, row_atual, row_idx, log_callback):
        super().__init__(parent)
        self.resultado    = resultado
        self.row_atual    = row_atual
        self.row_idx      = row_idx
        self.log_callback = log_callback
        self.confirmado   = False

        # Set de índices CSV já excluídos — cresce a cada "Ignorar e tentar novamente".
        # Começa com o match atual para que a primeira tentativa já o pule.
        self._excluidos = set()
        if row_atual.get('csv_idx') is not None:
            self._excluidos.add(row_atual['csv_idx'])

        self.title("✏️  Trocar Produto do CSV")
        self.geometry("820x330")
        self.configure(bg=self.COR_BG)
        self.grab_set()     # Modal — bloqueia interação com janela pai
        self.lift()
        self.focus_force()
        self.resizable(True, False)

        self._criar_widgets()

    def _criar_widgets(self):
        """
        Cria a interface do diálogo de troca:
          - Header com título
          - Info da linha atual (nome XLSX e match atual)
          - Combobox filtrável para escolher novo produto
          - Botões: Aplicar, Ignorar e tentar novamente, Cancelar
        """
        header = tk.Frame(self, bg=self.COR_HEADER, pady=8)
        header.pack(fill=tk.X)
        tk.Label(header, text="✏️  Trocar produto do CSV para esta linha",
                 font=("Segoe UI", 11, "bold"),
                 bg=self.COR_HEADER, fg="#c9a8ff").pack()

        # ── Informações da linha atual ────────────────────────────────
        frame_atual = tk.Frame(self, bg=self.COR_BG, padx=16, pady=10)
        frame_atual.pack(fill=tk.X)

        row = self.row_atual
        nome_xlsx = row.get('nome_xlsx') or row.get('nome_perfume') or '?'

        tk.Label(frame_atual, text="Nome XLSX:",
                 bg=self.COR_BG, fg="#888899",
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        tk.Label(frame_atual, text=nome_xlsx,
                 bg=self.COR_BG, fg="#66aaff",
                 font=("Consolas", 9, "bold")).grid(row=0, column=1, sticky="w", padx=8)

        tk.Label(frame_atual, text="Match atual:",
                 bg=self.COR_BG, fg="#888899",
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=(6,0))
        tk.Label(frame_atual,
                 text=f"SKU {row['sku']} | {row['descricao']}  "
                      f"(R$ {row['antes']} → R$ {row['depois']})",
                 bg=self.COR_BG, fg="#ffd700",
                 font=("Consolas", 9)).grid(row=1, column=1, sticky="w", padx=8, pady=(6,0))

        # ── Combobox filtrável para novo match ────────────────────────
        #  O combobox mostra TODOS os produtos do CSV no formato "[idx] Descrição".
        #  Ao digitar, o evento <KeyRelease> filtra as opções em tempo real.
        #  Isso é essencial porque pode haver centenas de produtos e
        #  scrollar seria muito lento.
        tk.Label(frame_atual, text="Novo produto:",
                 bg=self.COR_BG, fg="#888899",
                 font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=(8, 0))

        csv_descs = self.resultado.get('csv_descricoes', [])
        opcoes    = [f"[{idx}] {desc}" for idx, desc in csv_descs]
        # Mapa reverso: texto da opção → índice no CSV
        self._opcao_para_idx = {f"[{idx}] {desc}": idx for idx, desc in csv_descs}

        self._var = tk.StringVar()
        self._cb  = ttk.Combobox(frame_atual, textvariable=self._var,
                                  values=opcoes, width=70,
                                  font=("Consolas", 9))
        self._cb.grid(row=2, column=1, sticky="w", padx=8, pady=(8, 0))

        def _filtrar(event):
            """Filtra as opções do combobox conforme o usuário digita."""
            txt = self._cb.get().lower()
            self._cb['values'] = [o for o in opcoes if txt in o.lower()] or opcoes
        self._cb.bind("<KeyRelease>", _filtrar)

        # ── Botões ────────────────────────────────────────────────────
        frame_btn = tk.Frame(self, bg=self.COR_BG, pady=12)
        frame_btn.pack()

        tk.Button(frame_btn, text="✅  Aplicar",
                  bg="#15803d", fg="white",
                  font=("Arial", 10, "bold"), relief="flat",
                  cursor="hand2", padx=14, pady=5,
                  command=self._aplicar).pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(frame_btn, text="🔄  Ignorar match atual e tentar novamente",
                  bg="#92400e", fg="white",
                  font=("Arial", 10), relief="flat",
                  cursor="hand2", padx=14, pady=5,
                  command=self._ignorar_e_retentar).pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(frame_btn, text="Cancelar",
                  bg="#3a3a5c", fg="white",
                  font=("Arial", 10), relief="flat",
                  cursor="hand2", padx=14, pady=5,
                  command=self.destroy).pack(side=tk.LEFT)

    def _ignorar_e_retentar(self):
        """
        Remove o candidato atual do pool e roda o Jaccard novamente.

        Fluxo:
          1. Exclui TODOS os csv_idx já marcados em self._excluidos
          2. Recalcula palavras significativas para os CSVs restantes
          3. Roda o Jaccard para encontrar o próximo melhor candidato
          4. Verifica critérios de aceitação (cobertura + score)
          5. Se aceito → preenche o combobox com a sugestão
          6. Se rejeitado → avisa e permite tentar novamente ou escolher manual

        Pode ser clicado MÚLTIPLAS vezes — a cada clique o set _excluidos
        cresce e o algoritmo busca o próximo candidato na fila.
        """
        from logica_perfumaria import (_jaccard_score, palavras_sig,
                                        _palavras_fuzzy_match)

        resultado = self.resultado
        df_bling  = resultado['df_bling']
        col_desc  = resultado['col_desc']
        row       = self.row_atual
        nome_xlsx = row.get('nome_xlsx') or row.get('nome_perfume') or ''

        # Reconstrói o pool de candidatos excluindo os já ignorados
        csv_palavras_sem = {
            idx: palavras_sig(str(df_bling.at[idx, col_desc]))
            for idx in df_bling.index
            if idx not in self._excluidos
        }

        if not csv_palavras_sem:
            messagebox.showinfo(
                "Sem candidatos",
                "Todos os produtos do CSV já foram ignorados.\n"
                "Use o combobox para escolher manualmente.",
                parent=self
            )
            return

        # Roda o Jaccard no pool reduzido
        p_xlsx       = palavras_sig(nome_xlsx)
        SCORE_MINIMO = 0.35
        melhor_idx   = None
        melhor_score = 0.0

        for idx, p_csv in csv_palavras_sem.items():
            s = _jaccard_score(p_xlsx, p_csv)
            if s > melhor_score:
                melhor_score = s
                melhor_idx   = idx

        # Verifica cobertura total (mesma lógica da Passada 1)
        if melhor_idx is not None and p_xlsx:
            p_csv_m  = csv_palavras_sem[melhor_idx]
            casadas  = sum(1 for pw in p_xlsx
                           if any(_palavras_fuzzy_match(pw, pc) for pc in p_csv_m))
            cobertura = casadas / len(p_xlsx)
            todas     = cobertura == 1.0
        else:
            todas     = False
            cobertura = 0.0

        aceitar = melhor_idx is not None and todas and (
            melhor_score >= SCORE_MINIMO or cobertura == 1.0)

        if not aceitar:
            # Mesmo sem aceitar, adicionamos ao set para que a próxima
            # tentativa pule este candidato também
            if melhor_idx is not None:
                self._excluidos.add(melhor_idx)
            ignorados_n = len(self._excluidos)
            messagebox.showinfo(
                "Sem resultado adequado",
                f"Nenhum candidato válido além dos {ignorados_n} já ignorados.\n\n"
                f"Melhor score encontrado: {melhor_score:.2f} "
                f"(cobertura: {cobertura:.0%})\n\n"
                f"Clique novamente para continuar tentando,\n"
                f"ou use o combobox para escolher manualmente.",
                parent=self
            )
            return

        # Candidato aceito — adiciona ao set e preenche o combobox
        self._excluidos.add(melhor_idx)

        nova_desc = str(df_bling.at[melhor_idx, col_desc]).strip()
        opcao_str = f"[{melhor_idx}] {nova_desc}"
        self._var.set(opcao_str)
        self._opcao_para_idx[opcao_str] = melhor_idx

        ignorados_n = len(self._excluidos)
        messagebox.showinfo(
            f"Tentativa {ignorados_n} — Novo candidato",
            f"Sugestão (score {melhor_score:.2f}):\n\n"
            f"  {nova_desc}\n\n"
            f"• ✅ Aplicar  → confirma esta sugestão\n"
            f"• 🔄 Tentar novamente → pula este e busca o próximo\n"
            f"• Combobox → escolha manual livre",
            parent=self
        )

    def _aplicar(self):
        """
        Aplica a troca de match selecionada no combobox.

        Atualiza:
          1. O DataFrame df_bling (colunas de preço/estoque)
          2. O dicionário da linha em preview_rows
          3. O set indices_alterados (adiciona o novo csv_idx)

        A lógica de cálculo de preço é a mesma da Passada 2 e de
        aplicar_reassociacao — mantemos consistência replicando a fórmula.
        """
        escolha = self._var.get().strip()
        if not escolha or escolha not in self._opcao_para_idx:
            messagebox.showwarning("Atenção", "Selecione um produto da lista.", parent=self)
            return

        novo_csv_idx = self._opcao_para_idx[escolha]
        resultado    = self.resultado
        df_bling     = resultado['df_bling']
        col_desc     = resultado['col_desc']
        col_sku      = resultado.get('col_sku')
        col_gtin     = resultado.get('col_gtin')
        modo         = resultado['modo']

        # Lê dados do novo produto selecionado
        nova_desc = str(df_bling.at[novo_csv_idx, col_desc]).strip()
        novo_sku  = str(df_bling.at[novo_csv_idx, col_sku]).strip() if col_sku else 'N/A'
        novo_gtin = ''
        if col_gtin:
            val = str(df_bling.at[novo_csv_idx, col_gtin]).strip()
            if val.lower() not in ('nan', '', 'sem gtin', 'sem ean'):
                novo_gtin = val

        row = self.row_atual

        if modo == 'estoque':
            # Modo estoque: soma a quantidade ao estoque atual do novo produto
            est_ant = parse_moeda(str(df_bling.at[novo_csv_idx, 'Estoque']).strip()
                                  if 'Estoque' in df_bling.columns else '0')
            qtd     = parse_moeda(row['depois'])
            est_nov = est_ant + qtd
            fmt = lambda v: str(int(v)) if v == int(v) else f"{v:.2f}".replace('.', ',')
            df_bling.at[novo_csv_idx, 'Estoque'] = fmt(est_nov)
            row.update({
                'sku': novo_sku, 'descricao': nova_desc,
                'antes': fmt(est_ant), 'depois': fmt(est_nov),
                'gtin': novo_gtin, 'reassociado': True,
            })
        else:
            # Modo preço: recalcula preços para o novo produto
            cotacao = resultado['cotacao_dolar']
            frete   = resultado['taxa_frete']
            limite  = resultado['limite_variacao']

            # Recupera o valor unitário original a partir do preço de compra
            # que já foi calculado e está formatado em moeda
            val_unit= parse_moeda(row.get('compra', '0')) / cotacao if cotacao else parse_moeda(row.get('compra', '0'))
            preco_compra      = val_unit * cotacao
            preco_custo       = preco_compra * (1 + frete / 100.0)
            preco_venda_final = arredondar_preco(preco_custo * 2)
            preco_antigo      = str(df_bling.at[novo_csv_idx, 'Preço']).strip() \
                                if 'Preço' in df_bling.columns else '0,00'
            preco_novo        = formatar_moeda(preco_venda_final)
            diff_preco        = abs(preco_venda_final - parse_moeda(preco_antigo))
            nome_perfume      = row.get('desc_forn', nova_desc)
            qtd_ml_str        = row.get('desc_forn', '100ML')

            # Atualiza as colunas de preço no DataFrame para o novo produto
            if 'Preço de Compra' in df_bling.columns:
                df_bling.at[novo_csv_idx, 'Preço de Compra'] = formatar_moeda(preco_compra)
            if 'Preço de custo' in df_bling.columns:
                df_bling.at[novo_csv_idx, 'Preço de custo']  = formatar_moeda(preco_custo)
            if 'Preço' in df_bling.columns:
                df_bling.at[novo_csv_idx, 'Preço']            = preco_novo
            if 'Descrição do Produto no Fornecedor' in df_bling.columns:
                df_bling.at[novo_csv_idx, 'Descrição do Produto no Fornecedor'] = \
                    row.get('desc_forn', nova_desc)

            row.update({
                'sku': novo_sku, 'descricao': nova_desc,
                'antes': preco_antigo, 'depois': preco_novo,
                'gtin': novo_gtin, 'diff_preco': diff_preco,
                'variacao_pequena': (0 < diff_preco <= limite),
                'custo': formatar_moeda(preco_custo),
                'compra': formatar_moeda(preco_compra),
                'reassociado': True,
            })

        # Registra o novo csv_idx no set de alterados e atualiza preview_rows
        resultado['indices_alterados'].add(novo_csv_idx)
        resultado['preview_rows'][self.row_idx] = row

        self.log_callback(
            f"✏️  Match corrigido: SKU {novo_sku} | {nova_desc} "
            f"| R$ {row['antes']} → R$ {row['depois']}",
            "info"
        )
        self.confirmado = True
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
#  JANELA "VÁRIOS NUM SÓ"  (decomposição de produto composto)
# ──────────────────────────────────────────────────────────────────────────────
#
#  Quando a planilha do fornecedor tem um item como "ZAKAT VARIOS" com
#  Estoque = 8, mas na realidade são 4 perfumes diferentes (2 de cada),
#  esta janela permite decompor 1 item do XLSX em N produtos do CSV.
#
#  O usuário pode:
#    - Adicionar linhas (combo CSV + campo quantidade)
#    - Remover linhas individuais
#    - Validar que a soma das quantidades ≤ estoque total (modo estoque)
#    - Confirmar para gerar N reassociações independentes
#
#  Retorna uma lista de dicts, cada um com os dados para aplicar_reassociacao().
# ──────────────────────────────────────────────────────────────────────────────

class VariosNumSoWindow(tk.Toplevel):
    COR_BG     = "#12122a"
    COR_HEADER = "#1e1e3f"

    def __init__(self, parent, item_data, csv_descs, modo_estoque=False):
        """
        Janela para decompor um produto do XLSX em múltiplos do CSV.

        Parâmetros:
          parent        — janela pai (ReassociarWindow)
          item_data     — dict com dados do item XLSX (nome_perfume, quantidade_ml,
                          valor_unitario, qtd_estoque)
          csv_descs     — lista de tuplas (idx, desc) de todos os produtos do CSV
          modo_estoque  — se True, valida soma de quantidades ≤ estoque total
        """
        super().__init__(parent)
        self.item_data     = item_data
        self.csv_descs     = csv_descs
        self.modo_estoque  = modo_estoque
        self.confirmado    = False
        self.decomposicoes = []  # Resultado: lista de dicts {csv_idx, qtd_estoque, ...}

        # Mapa reverso: texto da opção → índice no CSV
        self._opcoes_csv = [f"[{idx}] {desc}" for idx, desc in csv_descs]
        self._opcao_para_idx = {f"[{idx}] {desc}": idx for idx, desc in csv_descs}

        self.title("📦  Vários num Só — Decomposição de Produto")
        self.geometry("920x520")
        self.configure(bg=self.COR_BG)
        self.grab_set()
        self.lift()
        self.focus_force()
        self.resizable(True, True)

        self._linhas = []  # Lista de dicts {frame, var_combo, var_qty}
        self._criar_widgets()

    def _criar_widgets(self):
        """
        Cria a interface:
          - Header com dados do item XLSX
          - Área scrollável de linhas (combo CSV + quantidade)
          - Botão "➕ Adicionar produto"
          - Barra de status com soma/total
          - Botões Aplicar/Cancelar
        """
        # ── Header com dados do item XLSX ──────────────────────────────
        header = tk.Frame(self, bg=self.COR_HEADER, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="📦  Vários num Só — Decompor Produto Composto",
                 font=("Helvetica", 12, "bold"),
                 bg=self.COR_HEADER, fg="#c9a8ff").pack()
        tk.Label(header,
                 text="Associe múltiplos produtos do CSV a este item do XLSX.",
                 font=("Segoe UI", 9),
                 bg=self.COR_HEADER, fg="#aaaacc").pack()

        # ── Info do item XLSX ─────────────────────────────────────────
        frame_info = tk.Frame(self, bg="#1a1a2e", padx=16, pady=10)
        frame_info.pack(fill=tk.X, padx=12, pady=(8, 4))

        item = self.item_data
        nome = item.get('nome_perfume', '?')
        ml   = item.get('quantidade_ml', 100)
        val  = item.get('valor_unitario', 0)
        est  = item.get('qtd_estoque', 0) or 0

        tk.Label(frame_info, text=f"XLSX:  {nome}",
                 bg="#1a1a2e", fg="#ffd700",
                 font=("Consolas", 10, "bold")).grid(row=0, column=0, sticky="w", columnspan=4)

        info_text = f"Volume: {int(ml) if ml == int(ml) else ml}ml   |   "
        info_text += f"Valor Unit.: US$ {val:.2f}   |   "
        info_text += f"Estoque Total: {int(est) if est == int(est) else est} un"
        tk.Label(frame_info, text=info_text,
                 bg="#1a1a2e", fg="#aaaacc",
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", columnspan=4, pady=(4, 0))

        # ── Botões de ação (Botton) ───────────────────────────────────
        frame_btn = tk.Frame(self, bg=self.COR_BG, pady=10)
        frame_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=12)
        tk.Button(frame_btn, text="✅  Aplicar Decomposição",
                  bg="#15803d", fg="white",
                  font=("Arial", 11, "bold"), relief="flat",
                  cursor="hand2", padx=16, pady=6,
                  command=self._aplicar).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(frame_btn, text="Cancelar",
                  bg="#3a3a5c", fg="white",
                  font=("Arial", 10), relief="flat",
                  cursor="hand2", padx=16, pady=6,
                  command=self.destroy).pack(side=tk.LEFT)

        # ── Botão Adicionar ───────────────────────────────────────────
        frame_add = tk.Frame(self, bg=self.COR_BG, pady=4)
        frame_add.pack(side=tk.BOTTOM, fill=tk.X, padx=12)
        tk.Button(frame_add, text="➕  Adicionar produto",
                  bg="#3a3a5c", fg="white",
                  font=("Segoe UI", 9), relief="flat",
                  cursor="hand2", padx=12, pady=3,
                  command=self._adicionar_linha).pack(side=tk.LEFT)
        tk.Button(frame_add, text="➗  Repartir Igualmente",
                  bg="#3a3a5c", fg="white",
                  font=("Segoe UI", 9), relief="flat",
                  cursor="hand2", padx=12, pady=3,
                  command=self._repartir_igualmente).pack(side=tk.LEFT, padx=(10, 0))

        # Label de status (soma / total)
        self._lbl_status = tk.Label(frame_add, text="",
                                     bg=self.COR_BG, fg="#aaaacc",
                                     font=("Segoe UI", 9))
        self._lbl_status.pack(side=tk.RIGHT, padx=8)

        # ── Área scrollável de linhas ─────────────────────────────────
        frame_scroll = tk.Frame(self, bg=self.COR_BG)
        # Empacotamos o scroll frame por último, para que ele ocupe apenas o espaço restante!
        frame_scroll.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=4)

        self._canvas = tk.Canvas(frame_scroll, bg=self.COR_BG, highlightthickness=0)
        sb = ttk.Scrollbar(frame_scroll, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._inner = tk.Frame(self._canvas, bg=self.COR_BG)
        self._cw = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(self._cw, width=e.width))

        # ── Cabeçalhos das colunas ────────────────────────────────────
        header_row = tk.Frame(self._inner, bg="#1e1e3f", pady=4, padx=8)
        header_row.pack(fill=tk.X)
        tk.Label(header_row, text="Produto do CSV",
                 bg="#1e1e3f", fg="#aaaaff",
                 font=("Segoe UI", 9, "bold"), width=58, anchor="w").grid(row=0, column=0, sticky="w")
        tk.Label(header_row, text="Qtd",
                 bg="#1e1e3f", fg="#aaaaff",
                 font=("Segoe UI", 9, "bold"), width=8, anchor="center").grid(row=0, column=1, padx=(8, 0))
        tk.Label(header_row, text="",
                 bg="#1e1e3f", width=4).grid(row=0, column=2)  # Espaço para botão remover

        # Adiciona a primeira linha automaticamente
        self._adicionar_linha()

        self._atualizar_status()

    def _adicionar_linha(self):
        """
        Adiciona uma nova linha (combo + campo quantidade + botão remover)
        à lista de decomposição.
        """
        i = len(self._linhas)
        bg_row = "#1a1a2e" if i % 2 == 0 else "#14142a"
        row_frame = tk.Frame(self._inner, bg=bg_row, pady=5, padx=8)
        row_frame.pack(fill=tk.X, pady=1)

        var_combo = tk.StringVar(value="")
        cb = ttk.Combobox(row_frame, textvariable=var_combo,
                          values=self._opcoes_csv, width=58,
                          font=("Consolas", 8))
        cb.grid(row=0, column=0, sticky="w")

        # Filtro em tempo real
        def _filtrar(event, cb=cb, opcoes=self._opcoes_csv):
            txt = cb.get().lower()
            cb['values'] = [o for o in opcoes if txt in o.lower()] or opcoes
        cb.bind("<KeyRelease>", _filtrar)

        var_qty = tk.StringVar(value="1")
        entry_qty = tk.Entry(row_frame, textvariable=var_qty,
                              width=8, font=("Consolas", 9),
                              justify="center")
        entry_qty.grid(row=0, column=1, padx=(8, 4))
        # Atualiza status quando a quantidade muda
        var_qty.trace_add("write", lambda *_: self._atualizar_status())

        # Botão remover linha
        linha_info = {'frame': row_frame, 'var_combo': var_combo, 'var_qty': var_qty}
        btn_rem = tk.Button(row_frame, text="🗑",
                            bg=bg_row, fg="#ff4d4d",
                            font=("Arial", 10), relief="flat",
                            cursor="hand2", padx=4,
                            command=lambda li=linha_info: self._remover_linha(li))
        btn_rem.grid(row=0, column=2, padx=(4, 0))

        self._linhas.append(linha_info)
        self._atualizar_status()

    def _repartir_igualmente(self):
        """
        Divide o estoque total igualmente entre as linhas atuais.
        Avisa o usuário se houver sobra (resto da divisão).
        """
        if not self.modo_estoque:
            return

        n_linhas = len(self._linhas)
        if n_linhas == 0:
            return

        try:
            est_total = float(self.item_data.get('qtd_estoque', 0) or 0)
            if est_total == int(est_total):
                est_total = int(est_total)
        except ValueError:
            est_total = 0

        if est_total <= 0:
            return

        qtd_por_linha = est_total // n_linhas
        resto = est_total % n_linhas

        for linha in self._linhas:
            # Converte para int se não tiver casas decimais, ou float
            linha['var_qty'].set(str(int(qtd_por_linha) if qtd_por_linha == int(qtd_por_linha) else qtd_por_linha))

        if resto > 0:
            messagebox.showinfo(
                "Resto na divisão",
                f"Foram distribuídas {qtd_por_linha} unidades para cada um dos {n_linhas} produtos.\n\n"
                f"Atenção: restaram {resto} unidade(s) de estoque!\n"
                f"Você deve adicioná-las manualmente nas quantidades acima.",
                parent=self
            )

    def _remover_linha(self, linha_info):
        """
        Remove uma linha da lista de decomposição.
        Mantém pelo menos 1 linha.
        """
        if len(self._linhas) <= 1:
            messagebox.showwarning("Atenção",
                                   "É necessário pelo menos 1 produto.",
                                   parent=self)
            return
        if linha_info in self._linhas:
            self._linhas.remove(linha_info)
            linha_info['frame'].destroy()
            self._atualizar_status()

    def _soma_quantidades(self):
        """Calcula a soma de todas as quantidades informadas."""
        total = 0
        for linha in self._linhas:
            try:
                qty = float(linha['var_qty'].get().replace(',', '.'))
                total += qty
            except (ValueError, TypeError):
                pass
        return total

    def _atualizar_status(self):
        """Atualiza o label de status com soma / estoque total."""
        soma = self._soma_quantidades()
        est_total = self.item_data.get('qtd_estoque', 0) or 0

        def fmt(v):
            return str(int(v)) if v == int(v) else f"{v:.1f}"

        if self.modo_estoque:
            cor = "#39ff14" if soma <= est_total else "#ff4d4d"
            texto = f"Soma: {fmt(soma)} / {fmt(est_total)} un"
        else:
            cor = "#aaaacc"
            texto = f"{len(self._linhas)} produto(s) | Soma qtd: {fmt(soma)}"

        self._lbl_status.config(text=texto, fg=cor)

    def _aplicar(self):
        """
        Valida e coleta as decomposições.

        Validações:
          - Pelo menos 1 linha preenchida
          - Nenhum combo vazio
          - Quantidades > 0
          - No modo estoque: soma ≤ estoque total
          - Sem csv_idx duplicados

        Resultado em self.decomposicoes: lista de dicts prontos para
        aplicar_reassociacao().
        """
        decomposicoes = []
        csv_idxs_usados = set()

        for i, linha in enumerate(self._linhas):
            escolha = linha['var_combo'].get().strip()
            qty_str = linha['var_qty'].get().strip().replace(',', '.')

            # Validação: combo não vazio
            if not escolha:
                messagebox.showwarning(
                    "Produto não selecionado",
                    f"Linha {i + 1}: selecione um produto do CSV.",
                    parent=self)
                return

            csv_idx = self._opcao_para_idx.get(escolha)
            if csv_idx is None:
                messagebox.showwarning(
                    "Produto inválido",
                    f"Linha {i + 1}: '{escolha}' não é um produto válido.",
                    parent=self)
                return

            # Validação: duplicatas
            if csv_idx in csv_idxs_usados:
                messagebox.showwarning(
                    "Produto duplicado",
                    f"Linha {i + 1}: este produto já foi adicionado.\n"
                    f"Use a quantidade para indicar mais de 1 unidade.",
                    parent=self)
                return
            csv_idxs_usados.add(csv_idx)

            # Validação: quantidade
            try:
                qty = float(qty_str)
                if qty <= 0:
                    raise ValueError
            except (ValueError, TypeError):
                messagebox.showwarning(
                    "Quantidade inválida",
                    f"Linha {i + 1}: '{linha['var_qty'].get()}' não é uma "
                    f"quantidade válida (deve ser > 0).",
                    parent=self)
                return

            decomposicoes.append({
                'csv_idx'       : csv_idx,
                'nome_perfume'  : self.item_data['nome_perfume'],
                'quantidade_ml' : self.item_data.get('quantidade_ml', 100),
                'valor_unitario': self.item_data.get('valor_unitario', 0),
                'qtd_estoque'   : qty,
            })

        if not decomposicoes:
            messagebox.showwarning("Atenção",
                                   "Adicione pelo menos 1 produto.",
                                   parent=self)
            return

        # Validação modo estoque: soma ≤ estoque total
        if self.modo_estoque:
            est_total = self.item_data.get('qtd_estoque', 0) or 0
            soma = sum(d['qtd_estoque'] for d in decomposicoes)
            if soma > est_total:
                messagebox.showwarning(
                    "Quantidade excedida",
                    f"A soma das quantidades ({soma:.0f}) excede o estoque "
                    f"total do XLSX ({est_total:.0f}).\n\n"
                    f"Ajuste as quantidades.",
                    parent=self)
                return

        self.decomposicoes = decomposicoes
        self.confirmado = True
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
#  JANELA DE REASSOCIAÇÃO MANUAL  (não encontrados automáticos)
# ──────────────────────────────────────────────────────────────────────────────
#
#  Quando o matching automático NÃO encontra um produto do XLSX no CSV,
#  o programa pergunta ao usuário se quer associá-los manualmente.
#
#  Esta janela mostra uma lista de todos os itens "não encontrados" e,
#  para cada um, um combobox filtrável com todos os produtos do CSV.
#  O usuário pode:
#    - Escolher o produto correto para cada item
#    - Deixar "(pular)" para ignorar o item
#    - Clicar "📦 Vários num Só" para decompor um item em múltiplos
#    - Clicar "Aplicar" para processar todos de uma vez
#
#  Os itens reassociados passam por aplicar_reassociacao() na logica,
#  que calcula preços e atualiza o DataFrame normalmente.
# ──────────────────────────────────────────────────────────────────────────────

class ReassociarWindow(tk.Toplevel):
    COR_BG     = "#12122a"
    COR_HEADER = "#1e1e3f"
    COR_ALERTA = "#ffd700"

    def __init__(self, parent, resultado, log_callback):
        super().__init__(parent)
        self.resultado    = resultado
        self.log_callback = log_callback
        self.confirmado   = False
        self.associacoes  = []      # Lista de dicts com as associações feitas
        self._varios_data = {}      # {indice_combo: [decomposicoes]} — itens decompostos via VariosNumSo

        self.title("🔍  Reassociação Manual — Não Encontrados")
        self.geometry("1040x500")
        self.configure(bg=self.COR_BG)
        self.grab_set()
        self.lift()
        self.focus_force()
        self.resizable(True, True)

        itens     = resultado.get('nao_encontrados_data', [])
        csv_descs = resultado.get('csv_descricoes', [])
        self._criar_widgets(itens, csv_descs)

    def _criar_widgets(self, itens, csv_descs):
        """
        Cria a interface de reassociação:
          - Header explicativo
          - Lista scrollável de itens não encontrados, cada um com combobox
          - Botões "Aplicar" e "Pular Tudo"

        Nota para o júnior: usamos Canvas + Frame interno para criar uma
        lista scrollável. Tkinter não tem um widget "scroll frame" nativo,
        então criamos manualmente:
          1. Canvas com scrollbar
          2. Frame interno posicionado como "window" dentro do Canvas
          3. bind <Configure> para atualizar scrollregion quando o conteúdo muda
        """
        header = tk.Frame(self, bg=self.COR_HEADER, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="🔍  Associação Manual de Produtos Não Encontrados",
                 font=("Helvetica", 12, "bold"),
                 bg=self.COR_HEADER, fg="#c9a8ff").pack()
        tk.Label(header,
                 text="Escolha o produto do CSV correspondente. Deixe '(pular)' para ignorar.  "
                      "Use '📦 Vários' se 1 item do XLSX contém múltiplos perfumes.",
                 font=("Segoe UI", 9),
                 bg=self.COR_HEADER, fg="#aaaacc").pack()

        # ── Área scrollável para a lista de itens ─────────────────────
        frame_scroll = tk.Frame(self, bg=self.COR_BG)
        frame_scroll.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        canvas = tk.Canvas(frame_scroll, bg=self.COR_BG, highlightthickness=0)
        sb     = ttk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame interno que vive DENTRO do Canvas e contém os widgets reais
        inner = tk.Frame(canvas, bg=self.COR_BG)
        cw    = canvas.create_window((0, 0), window=inner, anchor="nw")
        # Quando o conteúdo interno muda de tamanho, atualiza a scrollregion
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        # Quando o canvas é redimensionado, ajusta a largura do frame interno
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width))

        # ── Lista de comboboxes (1 por item não encontrado) ───────────
        opcoes_csv = ["(pular — não associar)"] + [
            f"[{idx}] {desc}" for idx, desc in csv_descs]
        self._opcao_para_idx = {f"[{idx}] {desc}": idx for idx, desc in csv_descs}
        self._combos = []   # Lista de tuplas (item_data, StringVar)

        for i, item in enumerate(itens):
            # Alterna cores de fundo para facilitar a leitura (zebra striping)
            bg_row = "#1a1a2e" if i % 2 == 0 else "#14142a"
            row    = tk.Frame(inner, bg=bg_row, pady=6, padx=8)
            row.pack(fill=tk.X, pady=1)

            # Label do nome do XLSX — guardamos referência para atualizar
            # quando o item for decomposto via VariosNumSo
            lbl_nome = tk.Label(row, text=f"XLSX: {item['nome_perfume']}",
                     bg=bg_row, fg=self.COR_ALERTA,
                     font=("Consolas", 9, "bold"),
                     width=38, anchor="w")
            lbl_nome.grid(row=0, column=0, sticky="w")
            tk.Label(row, text="→", bg=bg_row, fg="#888899",
                     font=("Arial", 12)).grid(row=0, column=1, padx=8)

            var = tk.StringVar(value="(pular — não associar)")
            cb  = ttk.Combobox(row, textvariable=var, values=opcoes_csv,
                               width=48, font=("Consolas", 8))
            cb.grid(row=0, column=2, sticky="w")

            # Filtro em tempo real: ao digitar, mostra apenas opções que contêm o texto
            def _filtrar(event, cb=cb, opcoes=opcoes_csv):
                txt = cb.get().lower()
                cb['values'] = [o for o in opcoes if txt in o.lower()] or opcoes
            cb.bind("<KeyRelease>", _filtrar)

            # Botão "📦 Vários num Só" — abre a VariosNumSoWindow para este item
            combo_idx = i  # Captura o índice para o closure
            btn_varios = tk.Button(
                row, text="📦 Vários",
                bg="#92400e", fg="white",
                font=("Segoe UI", 8, "bold"), relief="flat",
                cursor="hand2", padx=6, pady=2,
                command=lambda idx=combo_idx, it=item, v=var, cb_w=cb, lbl=lbl_nome:
                    self._abrir_varios_num_so(idx, it, v, cb_w, lbl)
            )
            btn_varios.grid(row=0, column=3, padx=(6, 0))

            self._combos.append((item, var))

        # ── Botões de ação ────────────────────────────────────────────
        frame_btn = tk.Frame(self, bg=self.COR_BG, pady=10)
        frame_btn.pack(fill=tk.X, padx=12)
        tk.Button(frame_btn, text="✅  Aplicar Associações",
                  bg="#15803d", fg="white",
                  font=("Arial", 11, "bold"), relief="flat",
                  cursor="hand2", padx=16, pady=6,
                  command=self._aplicar).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(frame_btn, text="⏭  Pular Tudo",
                  bg="#3a3a5c", fg="white",
                  font=("Arial", 10), relief="flat",
                  cursor="hand2", padx=16, pady=6,
                  command=self.destroy).pack(side=tk.LEFT)
        tk.Label(frame_btn, text="Itens pulados NÃO serão exportados.",
                 bg=self.COR_BG, fg="#666677",
                 font=("Segoe UI", 8)).pack(side=tk.RIGHT, padx=8)

    def _abrir_varios_num_so(self, combo_idx, item_data, var_combo, cb_widget, lbl_nome):
        """
        Abre a VariosNumSoWindow para decompor um item XLSX em múltiplos CSV.

        Se o usuário confirmar, guarda as decomposições em self._varios_data
        e atualiza visual do combo/label para indicar que o item foi decomposto.
        """
        modo_est = self.resultado.get('modo') == 'estoque'
        csv_descs = self.resultado.get('csv_descricoes', [])

        dlg = VariosNumSoWindow(self, item_data, csv_descs, modo_estoque=modo_est)
        self.wait_window(dlg)

        if dlg.confirmado and dlg.decomposicoes:
            # Guarda as decomposições associadas a este índice de combo
            self._varios_data[combo_idx] = dlg.decomposicoes

            # Atualiza visual: desabilita combo e mostra resumo
            n_prods = len(dlg.decomposicoes)
            soma_qty = sum(d.get('qtd_estoque', 0) or 0 for d in dlg.decomposicoes)
            resumo = f"📦 {n_prods} produtos ({int(soma_qty)} un)"
            var_combo.set(resumo)
            cb_widget.config(state="disabled")
            lbl_nome.config(fg="#66aaff")  # Azul para indicar que foi decomposto

    def _aplicar(self):
        """
        Coleta todas as associações feitas e fecha a janela.

        Para cada combobox que NÃO está em "(pular)", extrai o csv_idx
        e monta um dicionário com os dados necessários para
        aplicar_reassociacao() processar.

        Itens marcados como "vários num só" emitem MÚLTIPLAS associações
        (uma por produto CSV decomposto) em vez de uma só.

        A lista self.associacoes é lida pela PerfumariaApp após wait_window().
        """
        self.associacoes = []
        for combo_idx, (item, var) in enumerate(self._combos):
            # Verifica se este item foi decomposto via VariosNumSo
            if combo_idx in self._varios_data:
                # Emite N reassociações independentes
                self.associacoes.extend(self._varios_data[combo_idx])
                continue

            escolha = var.get()
            if escolha.startswith("(pular") or not escolha.strip():
                continue
            csv_idx = self._opcao_para_idx.get(escolha)
            if csv_idx is not None:
                # Mescla os dados do item (nome, ml, valor, estoque) com o csv_idx escolhido
                self.associacoes.append({**item, 'csv_idx': csv_idx})
        self.confirmado = True
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
#  JANELA PRINCIPAL
# ──────────────────────────────────────────────────────────────────────────────
#
#  Classe principal da aplicação. Gerencia:
#    - Campos de entrada (CSV, XLSX, cotação, frete, limite de variação)
#    - Checkbox do Modo Estoque
#    - Botão de processamento (com mudança visual conforme o modo)
#    - Log visual colorido (Text widget com tags de cor)
#    - Fluxo completo: processar → reassociar → filtrar variação → preview
#
#  O processamento roda em thread separada para não travar a UI.
#  A comunicação entre thread de processamento e UI é feita via queue.Queue.
# ──────────────────────────────────────────────────────────────────────────────

class PerfumariaApp:
    def __init__(self, root):
        """
        Inicializa a aplicação principal.

        Parâmetro root: instância de tk.Tk() — a janela raiz do Tkinter.
        Guardamos referência porque precisamos dela para:
          - Criar janelas filhas (Toplevel)
          - Usar after() para agendar callbacks na main thread
          - lift() e focus_force() para trazer janelas para frente
        """
        self.root = root
        self.root.title("AutoPreço Perfumaria — Atualizador Bling")
        self.root.geometry("800x720")
        self.root.resizable(True, True)
        self.root.configure(bg="#1e1e2e")

        # Variáveis Tkinter (StringVar/BooleanVar) são observáveis e se conectam
        # automaticamente aos widgets Entry/Checkbutton via parâmetro textvariable/variable.
        # Quando o usuário digita no Entry, a StringVar atualiza automaticamente e vice-versa.
        self.caminho_csv  = tk.StringVar()
        self.caminho_xlsx = tk.StringVar()
        self.cotacao_var  = tk.StringVar(value="")
        self.frete_var    = tk.StringVar(value="")
        self.limite_var   = tk.StringVar(value="10")   # Padrão: avisar se variação < R$10
        self.modo_estoque = tk.BooleanVar(value=False)

        self._criar_widgets()

    def _criar_widgets(self):
        """
        Cria toda a interface da janela principal.

        Layout vertical (de cima para baixo):
          1. Header com título e subtítulo
          2. Campos de arquivo (CSV e XLSX) com botões "Abrir"
          3. Parâmetros (cotação, frete, limite de variação)
          4. Checkbox Modo Estoque
          5. Separador
          6. Botão "PROCESSAR"
          7. Log de processamento (Text widget com scrollbar)

        Nota para o júnior: os dicionários EL e EE são "estilos" reutilizáveis.
        Em vez de repetir font=, bg=, fg= em cada widget, definimos uma vez
        e passamos como **kwargs. Isso segue o princípio DRY (Don't Repeat Yourself).
        """
        # ── Header ────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg="#12122a", pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="💎  AutoPreço Perfumaria",
                 font=("Helvetica", 17, "bold"),
                 bg="#12122a", fg="#c9a8ff").pack()
        tk.Label(header,
                 text="Atualização de Custos e Preços via Planilha de Fornecedor",
                 font=("Helvetica", 9),
                 bg="#12122a", fg="#888899").pack()

        # ── Campos de arquivo ─────────────────────────────────────────
        painel = tk.Frame(self.root, bg="#1e1e2e", padx=20, pady=10)
        painel.pack(fill=tk.X)
        EL = {"bg": "#1e1e2e", "fg": "#ddddee", "font": ("Segoe UI", 10)}   # Estilo para Labels
        EE = {"width": 56, "font": ("Consolas", 9)}                          # Estilo para Entries

        tk.Label(painel, text="CSV Bling:", **EL).grid(row=0, column=0, sticky="w", pady=4)
        tk.Entry(painel, textvariable=self.caminho_csv, **EE).grid(row=0, column=1, padx=6)
        tk.Button(painel, text="📂 Abrir", command=self.selecionar_csv,
                  bg="#3a3a5c", fg="white", relief="flat",
                  cursor="hand2").grid(row=0, column=2)

        tk.Label(painel, text="XLSX Fornecedor:", **EL).grid(row=1, column=0, sticky="w", pady=4)
        tk.Entry(painel, textvariable=self.caminho_xlsx, **EE).grid(row=1, column=1, padx=6)
        tk.Button(painel, text="📂 Abrir", command=self.selecionar_xlsx,
                  bg="#3a3a5c", fg="white", relief="flat",
                  cursor="hand2").grid(row=1, column=2)

        # ── Parâmetros de processamento ───────────────────────────────
        params = tk.Frame(self.root, bg="#1e1e2e", padx=20)
        params.pack(fill=tk.X)

        tk.Label(params, text="💵 Cotação do Dólar (R$):", **EL).grid(
            row=0, column=0, sticky="w", padx=(0, 6))
        self.entry_dolar = tk.Entry(params, textvariable=self.cotacao_var,
                                    width=12, font=("Consolas", 10))
        self.entry_dolar.grid(row=0, column=1, sticky="w")
        tk.Label(params, text="(deixe vazio se o valor já é em R$)",
                 bg="#1e1e2e", fg="#666677", font=("Segoe UI", 8)).grid(
            row=0, column=2, sticky="w", padx=10)

        tk.Label(params, text="🚚 % Frete:", **EL).grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=6)
        self.entry_frete = tk.Entry(params, textvariable=self.frete_var,
                                     width=12, font=("Consolas", 10))
        self.entry_frete.grid(row=1, column=1, sticky="w")
        tk.Label(params, text="(deixe vazio ou 0 para ignorar)",
                 bg="#1e1e2e", fg="#666677", font=("Segoe UI", 8)).grid(
            row=1, column=2, sticky="w", padx=10)

        tk.Label(params, text="🔔 Avisar variação < R$:", **EL).grid(
            row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        self.entry_limite = tk.Entry(params, textvariable=self.limite_var,
                                      width=12, font=("Consolas", 10))
        self.entry_limite.grid(row=2, column=1, sticky="w")
        tk.Label(params,
                 text="(pergunta individualmente para cada produto com variação pequena)",
                 bg="#1e1e2e", fg="#666677", font=("Segoe UI", 8)).grid(
            row=2, column=2, sticky="w", padx=10)

        # ── Checkbox Modo Estoque ─────────────────────────────────────
        #  Quando ativado:
        #    - Botão muda de cor (roxo → âmbar) e texto
        #    - Campos de cotação, frete e limite ficam desabilitados
        #    - Dica extra aparece ("O XLSX precisa ter a coluna Estoque (un)")
        frame_modo = tk.Frame(self.root, bg="#1a1a30", padx=20, pady=8)
        frame_modo.pack(fill=tk.X, padx=20, pady=(4, 0))
        self.chk_estoque = tk.Checkbutton(
            frame_modo,
            text="📦  Modo Estoque — atualiza SOMENTE a coluna Estoque (ignora preços)",
            variable=self.modo_estoque,
            command=self._toggle_modo_estoque,
            bg="#1a1a30", fg="#ffdd88", selectcolor="#1a1a30",
            activebackground="#1a1a30", activeforeground="#ffdd88",
            font=("Segoe UI", 10, "bold"), cursor="hand2"
        )
        self.chk_estoque.pack(side=tk.LEFT)
        self.lbl_modo_dica = tk.Label(
            frame_modo,
            text="O XLSX precisa ter a coluna 'Estoque (un)'",
            bg="#1a1a30", fg="#666677", font=("Segoe UI", 8)
        )
        self.lbl_modo_dica.pack(side=tk.LEFT, padx=12)
        # Esconde a dica inicialmente — só aparece quando modo estoque é ativado
        self.lbl_modo_dica.pack_forget()

        # ── Separador visual ─────────────────────────────────────────
        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=20, pady=8)

        # ── Botão PROCESSAR ──────────────────────────────────────────
        #  Cor e texto mudam conforme o modo (preço vs estoque).
        #  Durante processamento: desabilitado + texto "⏳ Processando..."
        self.btn_processar = tk.Button(
            self.root, text="🚀  PROCESSAR PERFUMARIA",
            bg="#7c3aed", fg="white",
            font=("Arial", 12, "bold"),
            activebackground="#6d28d9", activeforeground="white",
            relief="flat", cursor="hand2", pady=8,
            command=self.iniciar_processamento
        )
        self.btn_processar.pack(fill=tk.X, padx=20, pady=(0, 8))

        # ── Log de processamento ─────────────────────────────────────
        #  Um Text widget com tags de cor para diferenciar mensagens:
        #    sucesso (verde), erro (vermelho), alerta (amarelo), info (azul)
        #    gtin_amarelo, gtin_laranja, gtin_vermelho (escala GTIN)
        #
        #  O widget começa DISABLED para que o usuário não edite o log.
        #  Para escrever nele, habilitamos temporariamente, escrevemos, e
        #  desabilitamos novamente (ver método log()).
        frame_log = tk.Frame(self.root, bg="#1e1e2e", padx=20, pady=4)
        frame_log.pack(fill=tk.BOTH, expand=True)
        tk.Label(frame_log, text="📋 Log de Processamento",
                 bg="#1e1e2e", fg="#aaaacc",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

        container = tk.Frame(frame_log, bg="#0d0d1a")
        container.pack(fill=tk.BOTH, expand=True)
        self.text_log = tk.Text(container, bg="#0d0d1a", fg="#ccccdd",
                                font=("Consolas", 10), insertbackground="white",
                                relief="flat", pady=6, padx=8)
        sb = ttk.Scrollbar(container, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=sb.set)
        self.text_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Configura tags de cor para o log
        self.text_log.tag_configure("sucesso",       foreground="#39ff14")
        self.text_log.tag_configure("erro",          foreground="#ff4d4d")
        self.text_log.tag_configure("alerta",        foreground="#ffd700")
        self.text_log.tag_configure("info",          foreground="#66aaff")
        self.text_log.tag_configure("gtin_amarelo",  foreground="#ffd700")
        self.text_log.tag_configure("gtin_laranja",  foreground="#ff8c00")
        self.text_log.tag_configure("gtin_vermelho", foreground="#ff3b3b")
        self.text_log.config(state=tk.DISABLED)

        # Mensagens iniciais de orientação
        self.log("Sistema pronto. Selecione os arquivos e clique em Processar.", "info")
        self.log("Colunas esperadas no XLSX: 'Nome do Perfume' | 'Quantidade (ml)' | 'Valor Unitário'", "info")

    def log(self, mensagem, tipo=None):
        """
        Escreve uma mensagem no log visual com cor opcional.

        Nota: o Text widget fica DISABLED para impedir edição pelo usuário.
        Para escrever programaticamente, precisamos:
          1. Habilitar (NORMAL)
          2. Inserir texto com tag de cor
          3. Scrollar para o final (see(END))
          4. Desabilitar novamente (DISABLED)

        O update() força o Tkinter a redesenhar imediatamente,
        caso contrário as mensagens só apareceriam no final do processamento.
        """
        self.text_log.config(state=tk.NORMAL)
        self.text_log.insert(tk.END, mensagem + "\n", tipo or "")
        self.text_log.see(tk.END)
        self.text_log.config(state=tk.DISABLED)
        self.root.update()

    def _toggle_modo_estoque(self):
        """
        Callback do checkbox Modo Estoque.

        Quando ativado:
          - Desabilita campos de cotação, frete e limite (não fazem sentido para estoque)
          - Muda cor e texto do botão para âmbar "📦 PROCESSAR MODO ESTOQUE"
          - Mostra dica sobre a coluna "Estoque (un)"

        Quando desativado:
          - Reabilita tudo e volta ao visual padrão roxo
        """
        ativo  = self.modo_estoque.get()
        estado = tk.DISABLED if ativo else tk.NORMAL
        self.entry_dolar.config(state=estado)
        self.entry_frete.config(state=estado)
        self.entry_limite.config(state=estado)
        if ativo:
            self.btn_processar.config(text="📦  PROCESSAR MODO ESTOQUE", bg="#92400e")
            self.lbl_modo_dica.pack(side=tk.LEFT, padx=12)
        else:
            self.btn_processar.config(text="🚀  PROCESSAR PERFUMARIA", bg="#7c3aed")
            self.lbl_modo_dica.pack_forget()

    def selecionar_csv(self):
        """Abre file dialog para selecionar o CSV do Bling."""
        c = filedialog.askopenfilename(title="Selecione o CSV do Bling",
                                       filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if c:
            self.caminho_csv.set(c)
            self.log(f"CSV selecionado: {os.path.basename(c)}", "info")

    def selecionar_xlsx(self):
        """Abre file dialog para selecionar a planilha XLSX do fornecedor."""
        c = filedialog.askopenfilename(title="Selecione a Planilha do Fornecedor",
                                       filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if c:
            self.caminho_xlsx.set(c)
            self.log(f"XLSX selecionado: {os.path.basename(c)}", "info")

    def _validar_entradas(self):
        """
        Verifica se os campos obrigatórios foram preenchidos.
        Mostra messagebox de aviso se algum estiver vazio.
        Retorna True se tudo ok, False caso contrário.
        """
        if not self.caminho_csv.get():
            messagebox.showwarning("Atenção", "Selecione o arquivo CSV do Bling.", parent=self.root)
            return False
        if not self.caminho_xlsx.get():
            messagebox.showwarning("Atenção", "Selecione a planilha XLSX do fornecedor.", parent=self.root)
            return False
        return True

    def iniciar_processamento(self):
        """
        Inicia o processamento em thread separada.

        Fluxo:
          1. Valida entradas
          2. Limpa o log
          3. Parsea parâmetros numéricos (cotação, frete, limite)
          4. Desabilita o botão (impede duplo clique)
          5. Cria queue.Queue para comunicação thread-safe
          6. Inicia thread worker que chama processar_perfumaria()
          7. Inicia polling da fila via root.after()

        POR QUE THREAD?
          Tkinter é single-threaded. Se processar_perfumaria() rodar na
          main thread (que pode demorar segundos com centenas de produtos),
          a UI inteira TRAVA — o usuário não pode nem mover a janela.
          Com thread separada, a UI continua responsiva e as mensagens
          de log aparecem em tempo real.

        POR QUE QUEUE?
          Não podemos chamar self.log() da thread worker diretamente porque
          isso atualizaria widgets Tkinter fora da main thread (crash!).
          A thread enfileira mensagens, e a main thread drena a fila
          periodicamente via after().
        """
        if not self._validar_entradas():
            return

        # Limpa o log anterior
        self.text_log.config(state=tk.NORMAL)
        self.text_log.delete(1.0, tk.END)
        self.text_log.config(state=tk.DISABLED)

        # Parsea parâmetros
        modo_est        = self.modo_estoque.get()
        cotacao_dolar   = 1.0
        taxa_frete      = 0.0
        limite_variacao = 10.0

        if not modo_est:
            # Parsea cotação do dólar
            cotacao_str = self.cotacao_var.get().strip().replace(',', '.')
            if cotacao_str not in ("", "0"):
                try:
                    cotacao_dolar = float(cotacao_str)
                    self.log(f"Cotação do dólar: R$ {cotacao_dolar:.2f}", "info")
                except ValueError:
                    messagebox.showerror("Erro", f"Cotação inválida: '{cotacao_str}'", parent=self.root)
                    return
            else:
                self.log("Cotação não informada → valor já em R$.", "alerta")

            # Parsea percentual de frete
            frete_str = self.frete_var.get().strip().replace(',', '.')
            if frete_str:
                try:
                    taxa_frete = float(frete_str)
                    self.log(f"Frete: {taxa_frete:.1f}%", "info")
                except ValueError:
                    messagebox.showerror("Erro", f"Frete inválido: '{frete_str}'", parent=self.root)
                    return
            else:
                self.log("Frete não informado → 0%.", "alerta")

            # Parsea limite de variação
            lim_str = self.limite_var.get().strip().replace(',', '.')
            if lim_str:
                try:
                    limite_variacao = float(lim_str)
                except ValueError:
                    messagebox.showerror("Erro", f"Variação mínima inválida: '{lim_str}'", parent=self.root)
                    return
        else:
            self.log("Modo Estoque ativo — preços NÃO serão alterados.", "alerta")

        self.log("\nProcessando... aguarde.\n")

        # Desabilita o botão durante o processamento para impedir duplo clique
        self.btn_processar.config(state=tk.DISABLED, text="⏳  Processando...")

        # ── Fila thread-safe para mensagens de log ─────────────────────
        self._log_queue = queue.Queue()

        def _log_thread_safe(mensagem, tipo=None):
            """
            Chamado pela thread de processamento — enfileira para a main thread.
            NUNCA chame self.log() diretamente de uma thread secundária!
            """
            self._log_queue.put((mensagem, tipo))

        def _processar_fila_log():
            """
            Chamado pela main thread via after() — drena a fila de log.

            Funciona como um polling: a cada 100ms verifica se há mensagens
            na fila e as escreve no log. Continua rodando enquanto a thread
            de processamento estiver viva.
            """
            try:
                while True:
                    mensagem, tipo = self._log_queue.get_nowait()
                    self.log(mensagem, tipo)
            except queue.Empty:
                pass
            # Continua verificando enquanto a thread estiver viva
            if hasattr(self, '_thread_proc') and self._thread_proc.is_alive():
                self.root.after(100, _processar_fila_log)

        # ── Thread de processamento ────────────────────────────────────
        # Usamos lista [None] ao invés de variável simples porque threads
        # não podem atribuir a variáveis locais da função pai (escopo).
        # Com lista, acessamos self._resultado_holder[0] que é mutável.
        self._resultado_holder = [None]

        def _worker():
            """Thread worker que executa o processamento pesado."""
            try:
                res = processar_perfumaria(
                    arquivo_csv     = self.caminho_csv.get(),
                    arquivo_xlsx    = self.caminho_xlsx.get(),
                    cotacao_dolar   = cotacao_dolar,
                    taxa_frete      = taxa_frete,
                    log_callback    = _log_thread_safe,
                    modo_estoque    = modo_est,
                    limite_variacao = limite_variacao,
                )
                self._resultado_holder[0] = res
            except Exception as e:
                _log_thread_safe(f"[ERRO INESPERADO] {e}", "erro")
                self._resultado_holder[0] = None
            finally:
                # Agenda o pós-processamento na main thread.
                # after() é thread-safe no Tkinter — é o único método seguro
                # para agendar trabalho na main thread a partir de outra thread.
                self.root.after(50, lambda: self._pos_processamento(modo_est))

        self._thread_proc = threading.Thread(target=_worker, daemon=True)
        self._thread_proc.start()
        _processar_fila_log()  # Inicia o polling da fila de log

    def _pos_processamento(self, modo_est):
        """
        Executado na main thread após a thread de processamento terminar.

        Fluxo pós-processamento:
          1. Drena últimas mensagens de log da fila
          2. Restaura o botão de processamento
          3. Se houve erro fatal → para aqui
          4. Se há "não encontrados" → pergunta se quer reassociar manualmente
          5. Se há variações pequenas → pergunta individualmente por produto
          6. Abre janela de Preview para confirmação final
        """
        # Drena último lote de logs que podem ter chegado após o polling parar
        try:
            while True:
                mensagem, tipo = self._log_queue.get_nowait()
                self.log(mensagem, tipo)
        except queue.Empty:
            pass

        # Restaura o botão ao estado normal (texto e cor conforme o modo)
        if modo_est:
            self.btn_processar.config(state=tk.NORMAL, text="📦  PROCESSAR MODO ESTOQUE")
        else:
            self.btn_processar.config(state=tk.NORMAL, text="🚀  PROCESSAR PERFUMARIA")

        resultado = self._resultado_holder[0]

        if resultado is None:
            self.log("Erro fatal no processamento.", "erro")
            return

        # ── Reassociação manual para não encontrados ───────────────────
        #  Se algum produto do XLSX não foi encontrado automaticamente,
        #  pergunta ao usuário se quer associar manualmente.
        #  Isso abre a ReassociarWindow com comboboxes filtráveis.
        if resultado.get('nao_encontrados_data'):
            n = len(resultado['nao_encontrados_data'])
            self.root.lift()
            self.root.focus_force()
            if messagebox.askyesno("Produtos Não Encontrados",
                                   f"{n} produto(s) não foram localizados automaticamente.\n\n"
                                   f"Deseja associá-los manualmente ao CSV agora?",
                                   parent=self.root):
                janela_reassoc = ReassociarWindow(self.root, resultado, self.log)
                self.root.wait_window(janela_reassoc)
                if janela_reassoc.confirmado and janela_reassoc.associacoes:
                    aplicar_reassociacao(resultado, janela_reassoc.associacoes, self.log)
                    self.log(
                        f"✔ {len(janela_reassoc.associacoes)} item(s) reassociado(s).",
                        "alerta"
                    )

        if not resultado['preview_rows']:
            self.log("Nenhuma alteração para mostrar.", "erro")
            return

        # ── Pergunta individual para cada variação pequena ─────────────
        #  Se a diferença de preço é menor que o limite configurado,
        #  pergunta se deve aplicar o novo preço.
        #  "Sim" → exporta com preço novo.
        #  "Não" → reverte para o preço antigo, mas o produto AINDA É
        #          exportado (fica no indices_alterados).
        #  A única razão para um produto não ser exportado é não ter
        #  match no CSV (tratado na Passada 1).
        if not modo_est:
            df_bling = resultado['df_bling']
            for r in list(resultado['preview_rows']):
                if not r.get('variacao_pequena'):
                    continue
                aplicar = messagebox.askyesno(
                    "Variação Pequena — Confirmar?",
                    f"SKU {r['sku']}\n{r['descricao']}\n\n"
                    f"Preço atual : R$ {r['antes']}\n"
                    f"Preço novo  : R$ {r['depois']}\n"
                    f"Variação    : Δ R$ {r['diff_preco']:.2f}\n\n"
                    f"Deseja aplicar o novo preço?\n"
                    f"(O produto será exportado de qualquer forma.)",
                    parent=self.root
                )
                if not aplicar:
                    # Reverte o preço no DataFrame para o valor antigo,
                    # mas mantém o produto na exportação.
                    csv_idx = r.get('csv_idx')
                    if csv_idx is not None and 'Preço' in df_bling.columns:
                        df_bling.at[csv_idx, 'Preço'] = r['antes']
                    # Atualiza o preview para refletir que o preço não mudou
                    r['depois']         = r['antes']
                    r['diff_preco']     = 0.0
                    r['variacao_pequena'] = False
                    self.log(
                        f"ℹ️  SKU {r['sku']}: preço mantido em R$ {r['antes']} (variação recusada, produto exportado).",
                        "alerta"
                    )

        if not resultado['preview_rows']:
            self.log("Nenhum produto restante após filtros.", "erro")
            return

        # ── Preview ───────────────────────────────────────────────────
        #  Abre a janela de preview para o usuário conferir e confirmar.
        #  lift() e focus_force() garantem que a janela principal esteja
        #  visível antes de abrir o preview (evita bug do Windows onde
        #  diálogos abrem atrás da janela principal).
        self.root.lift()
        self.root.focus_force()
        preview = PreviewWindow(self.root, resultado, self.log)
        self.root.wait_window(preview)

        if preview.confirmado:
            self.log("\n✅ Exportação concluída com sucesso!", "sucesso")
        else:
            self.log("\n⚠️  Exportação cancelada pelo usuário.", "alerta")


# ── Entry point ───────────────────────────────────────────────────────────────
#
#  Ponto de entrada do programa. Cria a janela raiz do Tkinter,
#  instancia a aplicação e inicia o loop de eventos.
#
#  mainloop() é um loop infinito que processa eventos (cliques, teclas,
#  redraws, timers via after(), etc.) até a janela ser fechada.
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    root = tk.Tk()
    app  = PerfumariaApp(root)
    root.mainloop()
