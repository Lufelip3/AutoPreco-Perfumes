# 💎 AutoPreço Perfumes (v2.1)

Ferramenta desktop em **Python** para atualização automatizada de **preços** e **estoque** de perfumes no ERP [Bling](https://www.bling.com.br/).

Recebe uma planilha XLSX do fornecedor e um CSV exportado do Bling, cruza os dados por nome de perfume com **algoritmo de matching fuzzy (Jaccard)**, calcula os preços finais aplicando câmbio e frete, suporta **decomposição 1:N ("Vários num Só")**, e exporta um CSV pronto para reimportação no Bling.

> **🆕 Novidades da Versão 2.1**: Suporte completo a **produtos compostos (1:N)** com a janela *"Vários num Só"*, botão de divisão automática *"➗ Repartir Igualmente"* com alerta de sobras, e orquestração de múltiplas reassociações para um único item do fornecedor.

---

## 📐 Visão Geral da Arquitetura

O programa é dividido em **dois arquivos** com responsabilidades bem separadas:

```
interface_perfumaria.pyw  →  Tudo que o usuário VÊ e INTERAGE (Tkinter GUI, Popups, Modais)
logica_perfumaria.py      →  Tudo que o programa CALCULA e PROCESSA (Pandas, Jaccard, Preços)
```

A interface **nunca** faz cálculos de preço ou matching. A lógica **nunca** cria janelas ou exibe mensagens. A comunicação entre eles é feita via dicionários e callbacks.

### Por que `.pyw`?

No Windows, a extensão `.pyw` faz o Python rodar **sem abrir a janela do terminal** — só a interface gráfica aparece. Isso permite executar com duplo clique sem ver a tela preta do CMD.

### Threading e Comunicação

O Tkinter (biblioteca gráfica do Python) é **single-threaded**: se o processamento pesado rodar na thread principal, a interface **trava** até terminar. Para resolver isso:

1. O processamento roda numa **thread separada** (`threading.Thread`)
2. A thread envia mensagens de log via **`queue.Queue`** (fila thread-safe)
3. A thread principal drena a fila a cada 100ms via `root.after()` e atualiza o log visual

> **Regra de ouro**: NUNCA atualize widgets do Tkinter de uma thread secundária. Isso causa crashes aleatórios. Sempre use fila + `after()`.

---

## 🔄 Fluxo Completo de Processamento

O programa segue um pipeline linear com pontos de intervenção manual:

```
┌─────────────────┐     ┌─────────────────┐     ┌──────────────────────┐
│  1. LER XLSX    │ ──→ │  2. LER CSV     │ ──→ │  3. PASSADA 1        │
│  (Fornecedor)   │     │  (Bling)        │     │  Matching Jaccard    │
└─────────────────┘     └─────────────────┘     └──────────┬───────────┘
                                                           │
                        ┌──────────────────────────────────┘
                        ▼
              ┌─────────────────────┐     ┌──────────────────────────┐
              │  4. DETECÇÃO DE     │ ──→ │  5. PASSADA 2            │
              │  CONFLITOS (N:1)    │     │  Aplicar Preço/Estoque   │
              └─────────────────────┘     └──────────┬───────────────┘
                                                     │
                        ┌────────────────────────────┘
                        ▼
              ┌───────────────────────┐     ┌─────────────────────────┐
              │  6. REASSOCIAÇÃO      │ ──→ │  7. FILTRO DE VARIAÇÃO  │
              │  MANUAL (1:1 / 1:N)   │     │  (pergunta individual)  │
              └───────────────────────┘     └──────────┬──────────────┘
                                                       │
                        ┌──────────────────────────────┘
                        ▼
              ┌───────────────────────┐     ┌──────────────────────┐
              │  8. PREVIEW           │ ──→ │  9. EXPORTAR CSV     │
              │  (conferência visual) │     │  (perfumes_atualizados│
              └───────────────────────┘     │  .csv)               │
                                            └──────────────────────┘
```

### Detalhamento de cada etapa

#### 1–2. Leitura dos Arquivos

- **CSV do Bling**: lido com `pandas.read_csv()`, separador `;`, encoding `utf-8-sig` (ignora BOM)
- **XLSX do Fornecedor**: lido com `pandas.read_excel()`, `dtype=str` para evitar conversões numéricas automáticas
- O XLSX passa por **detecção flexível de colunas** — aceita variações de nome (ver seção Colunas Esperadas)

#### 3. Passada 1 — Matching Jaccard

Para cada perfume do XLSX, busca o melhor candidato no CSV (ver seção Algoritmo de Matching).

#### 4. Detecção de Conflitos

Se 2+ produtos do XLSX apontam para o **mesmo** produto do CSV → **conflito**.
Os itens conflitantes são **removidos do mapeamento automático** e enviados para a lista de "não encontrados", dando ao usuário a chance de resolvê-los manualmente na tela de reassociação.

> **Por que conflitos são perigosos?** Se dois perfumes do XLSX apontam para o mesmo SKU no CSV, o último processado "esmaga" o preço do primeiro silenciosamente. O programa bloqueia isso.

#### 5. Passada 2 — Aplicação em Memória

Os preços/estoques são calculados e escritos no DataFrame do Pandas **em memória** — nenhum arquivo é salvo nesta etapa. Isso permite que o usuário revise tudo antes de gravar.

#### 6. Reassociação Manual (Suporte 1:1 e 1:N)

Se há produtos não encontrados (incluindo os conflitantes), o programa abre a janela de Reassociação:
- Lista de todos os itens pendentes com combobox filtrável.
- **Associação 1:1**: selecione diretamente o produto correspondente do CSV.
- **Associação 1:N ("📦 Vários")**: para itens como `"ZAKAT VARIOS"` (composto por múltiplos perfumes), abra a janela de decomposição e associe N produtos individuais do CSV com suas quantidades.

#### 7. Filtro de Variação (só Modo Preço)

Para cada produto cuja variação de preço é menor que o limite configurado (padrão R$10), o programa pergunta individualmente: "Deseja atualizar?". Isso evita alterações insignificantes que podem desalinhar com comparadores de preço.

#### 8. Preview

Janela modal com tabela completa de todos os produtos que serão exportados (incluindo os decompostos de lote). Permite:
- **Conferir** cada match visualmente
- **Trocar** um match errado (duplo clique ou clique direito)
- **Editar** o valor diretamente (duplo clique na coluna "Preço Novo" / "Est. Novo")
- **Desmarcar** linhas com checkbox (☑/☐) para excluí-las da exportação
- **Remover** linhas via menu de contexto

#### 9. Exportação

Ao clicar "Confirmar e Exportar", o programa:
1. Verifica se há produtos não associados (bloqueia no Modo Preço)
2. Remove linhas desmarcadas (☐) do conjunto de exportação
3. Grava **somente** as linhas alteradas no arquivo `perfumes_atualizados.csv`

---

## 📸 Funcionalidades

### 🆕 Decomposição "Vários num Só" (Versão 2.1)
- **Associação 1:N**: Associa um único item do XLSX (ex: `"ZAKAT VARIOS"` com 8 unidades) a múltiplos perfumes individuais no CSV do Bling (ex: 4 perfumes diferentes, 2 unidades de cada).
- **Divisão Automática ("➗ Repartir Igualmente")**: Reparte a quantidade total do XLSX uniformemente entre os produtos adicionados.
- **Gestão de Sobras / Restos**: Se a divisão não for exata (ex: 9 unidades divididas entre 4 itens = 2 para cada), o sistema aplica a divisão inteira (2) e emite um alerta informando quantas unidades sobraram para alocação manual.
- **Validação Dinâmica de Estoque**: Valida em tempo real se a soma das quantidades atribuídas é menor ou igual ao estoque total do item do fornecedor.

### Modo Preço (padrão)
- Importa CSV do Bling + XLSX do fornecedor
- Cruza produtos por nome usando matching fuzzy inteligente
- Aplica cotação do dólar e percentual de frete
- Calcula preço de compra, custo e venda (arredondado para terminar em `,90`)
- Exporta CSV atualizado para reimportação no Bling

### Modo Estoque
- Atualiza **somente** a coluna Estoque do CSV
- Substitui a quantidade no processamento automático
- Soma quantidades na reassociação manual (incluindo itens decompostos)
- Requer coluna `Estoque (un)` no XLSX

### Matching Inteligente
- **Jaccard Scoring** fuzzy-aware com threshold configurável
- Remoção automática de prefixos de fragrância (`Eau de Parfum`, `EDP`, `PERF`, etc.)
- Matching por prefixo e por similaridade de palavras (threshold 78%)
- Detecção e resolução de conflitos (2+ XLSX disputando 1 CSV → enviados para reassociação)
- Cobertura total: aceita match quando 100% das palavras do XLSX estão no CSV
- Remoção de códigos alfanuméricos para evitar falsos positivos

### Preview e Conferência
- **Janela de Preview** modal antes de salvar — nada é gravado automaticamente
- Tabela com: SKU, Nome XLSX, Descrição CSV, Preço Atual → Novo, Custo, GTIN
- **Checkboxes** (☑/☐) por linha para marcar/desmarcar itens da exportação
- **Edição inline** de preço/estoque com duplo clique na coluna de valor
- Código de cores por variação de preço para produtos com GTIN/EAN:

  | Cor | Condição |
  |-----|----------|
  | Branco | Preço não mudou |
  | 🟡 Amarelo | `0 < Δ ≤ R$10` |
  | 🟠 Laranja | `R$10 < Δ < R$80` |
  | 🔴 Vermelho | `Δ ≥ R$80` |
  | 🔵 Azul | Corrigido manualmente (reassociado / decomposto) |

### Correção Manual
- **Trocar match** no Preview: duplo clique ou clique direito na linha
  - Combobox filtrável com todos os produtos do CSV
  - Botão "Ignorar e tentar novamente" — roda Jaccard excluindo candidatos já rejeitados
  - Suporta múltiplas tentativas por produto
- **Reassociação manual** de produtos não encontrados e conflitantes
  - Suporta associação simples (1:1) e decomposição em lote (1:N via botão **📦 Vários**)
  - Itens reassociados aparecem em azul no Preview com tag `REASSOC.`
  - Funciona em ambos os modos (Preço e Estoque)
- **Editar valor** diretamente no preview (duplo clique no preço/estoque novo)
- **Remover linha** da exportação via menu de contexto

### Filtro de Variação
- Campo configurável "Avisar se variação < R$" (padrão: R$10)
- Para cada produto que se enquadra, pergunta individualmente: "Deseja atualizar?"

### Validações e Travas de Segurança
- **Conflito N:1 no Preview**: se duas linhas marcadas apontam para o mesmo produto CSV → botão Confirmar desabilitado (fica cinza) até desmarcar uma
- **Produtos não associados** (Modo Preço): exportação é bloqueada, exibindo lista de itens pendentes. O botão fica clicável para o usuário ver quais são, mas a exportação não prossegue
- **Produtos não associados** (Modo Estoque): são simplesmente ignorados, sem bloquear
- **Remoção da lista fantasma**: quando um produto é reassociado com sucesso, ele é removido da lista de "não encontrados" — evitando bloqueio indevido

### Estabilidade
- Processamento em **thread separada** — UI nunca trava
- Comunicação via `queue.Queue` thread-safe
- Diálogos sempre visíveis (`parent=self.root` + `lift()` + `focus_force()`)
- Layouts responsivos com garantia de ancoragem de botões no rodapé (`side=tk.BOTTOM`)
- Tratamento robusto de `NaN` em células vazias do XLSX
- Auto-instalação de dependências na primeira execução

---

## 🗂 Estrutura do Projeto

```
AutoPreco Perfumes/
├── interface_perfumaria.pyw   # Interface gráfica (Tkinter) — ponto de entrada e janelas visuais
├── logica_perfumaria.py       # Toda a lógica de negócio: leitura, matching, cálculo, reambiguação
├── README.md                  # Documentação completa do projeto
├── .gitignore                 # Exclui .csv, .xlsx e arquivos temporários
└── prompt_transcrever_nota.md # Prompt auxiliar (referência)
```

| Arquivo | Responsabilidade |
|---------|-----------------|
| `logica_perfumaria.py` | Leitura de CSV/XLSX, normalização de texto, Jaccard scoring, mapeamento de colunas, detecção de conflitos, cálculo de preços, aplicação de reassociações simples e múltiplas, exportação do CSV final |
| `interface_perfumaria.pyw` | Janela principal (`PerfumariaApp`), preview modal (`PreviewWindow`), troca de match (`TrocarMatchDialog`), reassociação manual (`ReassociarWindow`) e decomposição em lote (`VariosNumSoWindow`) |

---

## 🚀 Como Usar

### Pré-requisitos

- **Python 3.8+** instalado
- As dependências (`pandas`, `openpyxl`) são instaladas automaticamente na primeira execução

### Execução

```bash
# Duplo clique no arquivo .pyw (Windows) ou:
python interface_perfumaria.pyw
```

### Passo a Passo — Processamento Padrão

1. **Selecionar CSV do Bling** — exportado com as colunas de produtos
2. **Selecionar XLSX do Fornecedor** — planilha com nomes, quantidades e valores
3. **Preencher parâmetros** (modo preço):
   - 💵 Cotação do dólar (deixe vazio se valores já estão em R$)
   - 🚚 Percentual de frete (deixe vazio ou 0 para ignorar)
   - 🔔 Limite de variação para alerta individual
4. **Clicar em "Processar"**
5. **Resolver não encontrados**:
   - Para itens 1:1, escolha o produto no combobox.
   - Para itens compostos (ex: `"ZAKAT VARIOS"`), clique em **"📦 Vários"**, adicione os perfumes, clique em **"➗ Repartir Igualmente"** e confirme.
6. **Responder filtro de variação** — para variações pequenas, confirmar produto a produto
7. **Revisar o Preview** — conferir cada linha, trocar matches errados, editar valores, marcar/desmarcar linhas
8. **Confirmar e Exportar** — gera `perfumes_atualizados.csv` na mesma pasta do CSV original

---

## 📋 Colunas Esperadas

### XLSX do Fornecedor

O sistema aceita variações de nome para cada coluna (mapeamento flexível):

| Coluna Interna | Apelidos Aceitos |
|----------------|-----------------|
| `Nome do Perfume` | `nome formatado (erp)`, `produto`, `nome`, `perfume`, `description` |
| `Quantidade (ml)` | `quantidade`, `ml`, `volume`, `qty` |
| `Valor Unitario` | `valor unitário`, `valor unitario (us$)`, `unit price`, `price`, `preco` |
| `Estoque (un)` *(modo estoque)* | `estoque`, `quantidade (un)`, `unidades`, `stock`, `qtd estoque` |

A comparação de nomes de colunas é **case-insensitive** e **sem acentos** — `"Valor Unitário (US$)"` bate com `"valor unitario (us$)"`.

### CSV do Bling

Colunas utilizadas (nomes padrão do Bling):
- `Descrição` — nome do produto
- `Código` — SKU
- `Preço` — preço de venda
- `Preço de Compra` — preço de compra
- `Preço de custo` — preço de custo
- `GTIN/EAN` — código de barras (para alertas)
- `Estoque` — quantidade em estoque (modo estoque)
- `Descrição do Produto no Fornecedor` — descrição formatada

---

## ⚙️ Lógica de Preços

### Fórmula

```
Preço de Compra = Valor Unitário (XLSX) × Cotação do Dólar
Preço de Custo  = Preço de Compra × (1 + Frete%)
Preço de Venda  = arredonda(Preço de Custo × 2)  → termina em ,90
```

### Arredondamento

O preço final sempre termina em **,90** (precificação psicológica):
```
arredondar_preco(valor) = round(valor / 10) * 10 - 0.10
```

---

## 🧠 Algoritmo de Matching

O matching entre XLSX (fornecedor) e CSV (Bling) funciona em **2 passadas**:

### Passada 1 — Mapeamento Jaccard

1. **Remove prefixos de fragrância** (`Eau de Parfum`, `EDP`, `PERF`, `Cologne`, etc.)
2. **Normaliza texto** (lowercase, remove acentos via decomposição Unicode NFKD)
3. **Remove códigos numéricos** (tokens que começam com dígito: `2230`, `100M6173`)
4. **Filtra stopwords** do domínio de perfumaria (`de`, `da`, `eau`, `ml`, `parfum`, etc.)
5. **Extrai palavras significativas** — resultado é um SET (ordem não importa)
6. **Calcula Jaccard Score fuzzy** contra todos os produtos do CSV (threshold 78% / prefixos)

### Passada 2 — Detecção e Resolução de Conflitos

- Se 2+ produtos do XLSX apontam para o mesmo CSV → **conflito N:1**
- Os itens conflitantes são enviados para a tela de Reassociação Manual
- Permite associação 1:1 direta ou decomposição 1:N via **"Vários num Só"**

---

## 🖥 Interface Gráfica — Classes e Janelas

### `PerfumariaApp` — Janela Principal
A classe principal que gerencia toda a aplicação, orquestração de threads e logs.

### `PreviewWindow` — Janela de Preview
Tabela modal com todos os produtos prontos para exportação. Suporta seleção com checkboxes, edição inline e menu de contexto.

### `TrocarMatchDialog` — Trocar Match
Permite corrigir um match errado do algoritmo individualmente no preview.

### `ReassociarWindow` — Reassociação Manual
Exibe lista de itens não encontrados / conflitantes. Apresenta o combobox individual e o botão **"📦 Vários"** por linha.

### 🆕 `VariosNumSoWindow` — Decomposição de Lote (v2.1)
Janela modal dedicada a desmembrar um item do XLSX em múltiplos produtos do CSV:
- Adição/Remoção dinâmica de linhas (combo do CSV + campo de quantidade).
- Botão **"➗ Repartir Igualmente"** com cálculo de restos e pop-up informativo.
- Barra de status com totalizadores e validação contínua de estoque.
- Empacotamento de layout corrigido no rodapé (`side=tk.BOTTOM`).

---

## 🔖 Histórico de Versões

### **v2.1** (Atual)
- ✨ **Suporte a Associação 1:N ("Vários num Só")**: Permite decompor um item do XLSX em múltiplos produtos do CSV do Bling.
- ➗ **Divisão Automática de Quantidades**: Novo botão "Repartir Igualmente" com tratamento inteligente de sobras/restos de divisão.
- 🎨 **Correção de Layout em Telas Menores**: Reestruturação do empacotamento (`pack`) em modais para garantir visibilidade permanente dos botões no rodapé.
- ⚡ **Orquestração de Reassociações Múltiplas**: Atualização da ponte de dados entre GUI e Engine de Cálculo para emitir N atualizações a partir de 1 registro.

### **v2.0**
- 📦 Suporte completo ao **Modo Estoque**.
- 🔍 Visualização avançada em **Dark Mode** com temas customizados.
- 🛡 Travas de segurança para conflitos N:1 e verificação de itens não associados.

---

## 📝 Licença

Projeto interno de uso pessoal / comercial privado.
