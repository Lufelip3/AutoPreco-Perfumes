# 💎 AutoPreço Perfumes

Ferramenta desktop em **Python** para atualização automatizada de **preços** e **estoque** de perfumes no ERP [Bling](https://www.bling.com.br/).

Recebe uma planilha XLSX do fornecedor e um CSV exportado do Bling, cruza os dados por nome de perfume com **algoritmo de matching fuzzy (Jaccard)**, calcula os preços finais aplicando câmbio e frete, e exporta um CSV pronto para reimportação no Bling.

---

## 📐 Visão Geral da Arquitetura

O programa é dividido em **dois arquivos** com responsabilidades bem separadas:

```
interface_perfumaria.pyw  →  Tudo que o usuário VÊ e INTERAGE
logica_perfumaria.py      →  Tudo que o programa CALCULA e PROCESSA
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
              │  MANUAL (janela)      │     │  (pergunta individual)  │
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

#### 6. Reassociação Manual

Se há produtos não encontrados (incluindo os conflitantes), o programa abre uma janela com:
- Lista de todos os itens pendentes
- Um combobox filtrável para cada item, com todos os produtos do CSV
- Opção de "pular" itens que não precisam ser associados

A reassociação funciona **tanto no Modo Preço quanto no Modo Estoque**.

#### 7. Filtro de Variação (só Modo Preço)

Para cada produto cuja variação de preço é menor que o limite configurado (padrão R$10), o programa pergunta individualmente: "Deseja atualizar?". Isso evita alterações insignificantes que podem desalinhar com comparadores de preço.

#### 8. Preview

Janela modal com tabela completa de todos os produtos que serão exportados. Permite:
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

### Modo Preço (padrão)
- Importa CSV do Bling + XLSX do fornecedor
- Cruza produtos por nome usando matching fuzzy inteligente
- Aplica cotação do dólar e percentual de frete
- Calcula preço de compra, custo e venda (arredondado para terminar em `,90`)
- Exporta CSV atualizado para reimportação no Bling

### Modo Estoque
- Atualiza **somente** a coluna Estoque do CSV
- Substitui a quantidade no processamento automático
- Soma quantidades na reassociação manual (não substitui)
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
  | 🔵 Azul | Corrigido manualmente (reassociado) |

### Correção Manual
- **Trocar match** no Preview: duplo clique ou clique direito na linha
  - Combobox filtrável com todos os produtos do CSV
  - Botão "Ignorar e tentar novamente" — roda Jaccard excluindo candidatos já rejeitados
  - Suporta múltiplas tentativas por produto
- **Reassociação manual** de produtos não encontrados e conflitantes
  - Janela dedicada com combobox filtrável para cada item
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
- Tratamento robusto de `NaN` em células vazias do XLSX
- Auto-instalação de dependências na primeira execução

---

## 🗂 Estrutura do Projeto

```
AutoPreco Perfumes/
├── interface_perfumaria.pyw   # Interface gráfica (Tkinter) — ponto de entrada
├── logica_perfumaria.py       # Toda a lógica de negócio: leitura, matching, cálculo
├── README.md                  # Este arquivo
├── .gitignore                 # Exclui .csv, .xlsx e arquivos temporários
└── prompt_transcrever_nota.md # Prompt auxiliar (referência)
```

| Arquivo | Responsabilidade |
|---------|-----------------|
| `logica_perfumaria.py` | Leitura de CSV/XLSX, normalização de texto, Jaccard scoring, mapeamento de colunas, detecção de conflitos, cálculo de preços, reassociação, exportação do CSV final |
| `interface_perfumaria.pyw` | Janela principal, campos de entrada, log visual colorido, preview modal, troca de match inline, edição de valor, checkbox de seleção, reassociação manual, menu de contexto |

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

### Passo a passo

1. **Selecionar CSV do Bling** — exportado com as colunas de produtos
2. **Selecionar XLSX do Fornecedor** — planilha com nomes, quantidades e valores
3. **Preencher parâmetros** (modo preço):
   - 💵 Cotação do dólar (deixe vazio se valores já estão em R$)
   - 🚚 Percentual de frete (deixe vazio ou 0 para ignorar)
   - 🔔 Limite de variação para alerta individual
4. **Clicar em "Processar"**
5. **Resolver não encontrados** — se houver, a janela de reassociação abre automaticamente
6. **Responder filtro de variação** — para variações pequenas, confirmar produto a produto
7. **Revisar o Preview** — conferir cada linha, trocar matches errados, editar valores, marcar/desmarcar linhas
8. **Confirmar e Exportar** — gera `perfumes_atualizados.csv` na mesma pasta do CSV original

### Modo Estoque

1. Marcar checkbox **"📦 Modo Estoque"**
2. Campos de cotação, frete e limite ficam desabilitados
3. O XLSX precisa ter a coluna `Estoque (un)`
4. Processar → Reassociar (se necessário) → Preview → Confirmar

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

### Exemplo completo

| Etapa | Cálculo | Resultado |
|-------|---------|-----------|
| Valor unitário (XLSX) | — | US$ 15.00 |
| Cotação do dólar | — | R$ 5.50 |
| Frete | — | 10% |
| **Preço de Compra** | 15.00 × 5.50 | R$ 82,50 |
| **Preço de Custo** | 82.50 × 1.10 | R$ 90,75 |
| **Preço de Venda** | arredonda(90.75 × 2) = arredonda(181.50) | **R$ 179,90** |

### Descrição do Fornecedor

Para cada produto processado, o programa gera automaticamente uma string padronizada para o campo "Descrição do Produto no Fornecedor" do Bling:

```
NOME DO PERFUME 100ML 3.4 FL.OZ.
```

Se a quantidade em ml não estiver disponível, assume 100ml como fallback.

---

## 🧠 Algoritmo de Matching

O matching entre XLSX (fornecedor) e CSV (Bling) funciona em **2 passadas**:

### Passada 1 — Mapeamento Jaccard

Para cada perfume do XLSX:

1. **Remove prefixos de fragrância** (`Eau de Parfum`, `EDP`, `PERF`, `Cologne`, etc.)
2. **Normaliza texto** (lowercase, remove acentos via decomposição Unicode NFKD)
3. **Remove códigos numéricos** (tokens que começam com dígito: `2230`, `100M6173`)
4. **Filtra stopwords** do domínio de perfumaria (`de`, `da`, `eau`, `ml`, `parfum`, etc.)
5. **Extrai palavras significativas** — resultado é um SET (ordem não importa)
6. **Calcula Jaccard Score fuzzy** contra todos os produtos do CSV

#### Jaccard Fuzzy

O Jaccard clássico é `|A ∩ B| / |A ∪ B|`, mas aqui usamos uma versão "fuzzy": duas palavras são consideradas par se:

- Uma é **prefixo** da outra (fornecedores abreviam: `"Marshm"` → `"Marshmallow"`)
- A **similaridade** (SequenceMatcher) é ≥ 78% (`"Musaman"` vs `"Musamam"` → ~85%)

#### Critérios de Aceitação

Um match é aceito se:
1. **TODAS** as palavras do XLSX têm ao menos 1 par fuzzy no CSV — evita `"Aseel Maison Asrar"` bater em `"Rey Maison Asrar"` só porque `"Maison"` + `"Asrar"` casam
2. **Score ≥ 0.35** OU **cobertura 100%** — a regra de cobertura resolve produtos de 1 palavra (ex: `"Aabinus"` com Jaccard 0.33, abaixo do mínimo, mas 100% de cobertura)

#### Stopwords e Palavras Distintivas

Palavras como `"intense"`, `"masc"` **NÃO** são ignoradas de propósito:
- `"Club Nuit Intense"` é um produto DIFERENTE de `"Club Nuit"`
- Remover essas palavras causaria falsos positivos

### Passada 2 — Detecção e Resolução de Conflitos

- Se 2+ produtos do XLSX apontam para o mesmo CSV → **conflito**
- Os itens conflitantes são **removidos** do mapeamento automático
- São **adicionados** à lista de "não encontrados" para reassociação manual
- O usuário resolve via dropdown na janela de reassociação, apontando cada um para o produto correto

---

## 🖥 Interface Gráfica — Classes e Janelas

### `PerfumariaApp` — Janela Principal

A classe principal que gerencia toda a aplicação:

- **Campos de entrada**: CSV, XLSX, cotação do dólar, frete, limite de variação
- **Checkbox Modo Estoque**: alterna visual (roxo → âmbar) e desabilita campos de preço
- **Botão Processar**: inicia thread de processamento, muda para "⏳ Processando..."
- **Log visual**: Text widget com tags de cor (verde=sucesso, vermelho=erro, amarelo=alerta, azul=info)
- **Orquestração pós-processamento**: reassociação → filtro variação → preview

### `PreviewWindow` — Janela de Preview

Tabela modal com todos os produtos prontos para exportação:

- **Treeview** estilizado em dark mode com scrollbars
- **Checkbox por linha** (☑/☐) para marcar/desmarcar
- **Duplo clique** na coluna de valor → edição inline
- **Duplo clique** em outras colunas → abre `TrocarMatchDialog`
- **Clique direito** → menu de contexto (trocar match ou remover)
- **Validação N:1**: desabilita botão Confirmar se há duplicatas marcadas
- **Bloqueio de exportação**: no Modo Preço, impede exportar se há produtos não associados

### `TrocarMatchDialog` — Trocar Match

Permite corrigir um match errado do algoritmo:

- Mostra nome do XLSX e match atual
- Combobox filtrável com todos os produtos do CSV
- Botão "Ignorar e tentar novamente" → roda Jaccard excluindo candidatos já rejeitados
- Suporta **múltiplas tentativas** (set de excluídos cresce a cada tentativa)

### `ReassociarWindow` — Reassociação Manual

Para produtos não encontrados e conflitantes:

- Lista scrollável com canvas + frame interno (scroll frame manual do Tkinter)
- Um combobox filtrável por item, com opção "(pular)"
- Zebra striping para facilitar leitura
- Funciona em Modo Preço e Modo Estoque

---

## 🛡 Mecanismos de Segurança

| Mecanismo | O que evita |
|-----------|-------------|
| Processamento em memória | Nenhum arquivo é gravado até o usuário confirmar no Preview |
| Detecção de conflitos N:1 | Impede que dois perfumes sobreescrevam o mesmo SKU silenciosamente |
| Bloqueio de exportação (não associados) | No Modo Preço, impede exportar se há itens pendentes |
| Remoção da lista fantasma | Ao reassociar, remove o item da lista de bloqueio — evita bloqueio indevido |
| Filtro de variação | Pergunta individualmente para preços com variação pequena |
| Checkbox de seleção | Permite excluir linhas específicas sem remover do preview |
| Exportação parcial | Grava somente as linhas alteradas, não o CSV inteiro |
| Thread separada + Queue | Evita travar a UI durante processamento pesado |
| Auto-instalação de dependências | Verifica e instala `pandas` e `openpyxl` automaticamente |

---

## 🛠 Tecnologias

- **Python 3.8+**
- **Tkinter** — interface gráfica nativa
- **Pandas** — manipulação de dados CSV/XLSX
- **openpyxl** — leitura de arquivos Excel (.xlsx)
- **difflib** — similaridade de strings (SequenceMatcher) para matching fuzzy
- **threading + queue** — processamento assíncrono sem travar a UI
- **unicodedata** — normalização de acentos (NFKD)
- **re** — remoção de prefixos de fragrância e códigos numéricos

---

## 📝 Licença

Projeto interno de uso pessoal / comercial privado.
