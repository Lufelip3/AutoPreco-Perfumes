# 💎 AutoPreço Perfumes

Ferramenta desktop em **Python** para atualização automatizada de **preços** e **estoque** de perfumes no ERP [Bling](https://www.bling.com.br/).

Recebe uma planilha XLSX do fornecedor e um CSV exportado do Bling, cruza os dados por nome de perfume com **algoritmo de matching fuzzy (Jaccard)**, calcula os preços finais aplicando câmbio e frete, e exporta um CSV pronto para reimportação no Bling.

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
- Soma quantidades (não substitui)
- Requer coluna `Estoque (un)` no XLSX

### Matching Inteligente
- **Jaccard Scoring** fuzzy-aware com threshold configurável
- Remoção automática de prefixos de fragrância (`Eau de Parfum`, `EDP`, `PERF`, etc.)
- Matching por prefixo e por similaridade de palavras (threshold 78%)
- Detecção e bloqueio de colisões (2+ XLSX disputando 1 CSV)
- Cobertura total: aceita match quando 100% das palavras do XLSX estão no CSV
- Remoção de códigos alfanuméricos para evitar falsos positivos

### Preview e Conferência
- **Janela de Preview** modal antes de salvar — nada é gravado automaticamente
- Tabela com: SKU, Nome XLSX, Descrição CSV, Preço Atual → Novo, Custo, GTIN
- Código de cores por variação de preço para produtos com GTIN/EAN:

  | Cor | Condição |
  |-----|----------|
  | Branco | Preço não mudou |
  | 🟡 Amarelo | `0 < Δ ≤ R$10` |
  | 🟠 Laranja | `R$10 < Δ < R$80` |
  | 🔴 Vermelho | `Δ ≥ R$80` |

### Correção Manual
- **Trocar match** no Preview: duplo clique ou clique direito na linha
  - Combobox filtrável com todos os produtos do CSV
  - Botão "Ignorar e tentar novamente" — roda Jaccard excluindo candidatos já rejeitados
  - Suporta múltiplas tentativas por produto
- **Reassociação manual** de produtos não encontrados
  - Janela dedicada com combobox filtrável para cada item
  - Itens reassociados aparecem em azul no Preview com tag `REASSOC.`
- **Remover linha** da exportação via menu de contexto

### Filtro de Variação
- Campo configurável "Avisar se variação < R$" (padrão: R$10)
- Para cada produto que se enquadra, pergunta individualmente: "Deseja atualizar?"

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
| `logica_perfumaria.py` | Leitura de CSV/XLSX, normalização, Jaccard scoring, mapeamento, cálculo de preços, exportação |
| `interface_perfumaria.pyw` | Janela principal, preview, reassociação manual, troca de match, log visual |

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
5. **Revisar o Preview** — conferir cada linha, trocar matches errados se necessário
6. **Confirmar e Exportar** — gera `perfumes_atualizados.csv` na mesma pasta do CSV original

### Modo Estoque

1. Marcar checkbox **"📦 Modo Estoque"**
2. Campos de cotação e frete ficam desabilitados
3. O XLSX precisa ter a coluna `Estoque (un)`
4. Processar → Preview → Confirmar

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

```
Preço de Compra = Valor Unitário (XLSX) × Cotação do Dólar
Preço de Custo  = Preço de Compra × (1 + Frete%)
Preço de Venda  = arredonda(Preço de Custo × 2)  → termina em ,90
```

**Exemplo:**
- Valor unitário: US$ 15.00
- Cotação: R$ 5.50
- Frete: 10%
- Compra: 15 × 5.50 = R$ 82.50
- Custo: 82.50 × 1.10 = R$ 90.75
- Venda: arredonda(90.75 × 2) = R$ 180 → R$ 179,90

---

## 🧠 Algoritmo de Matching

O matching entre XLSX (fornecedor) e CSV (Bling) funciona em **2 passadas**:

### Passada 1 — Mapeamento Jaccard
Para cada perfume do XLSX:
1. Remove prefixos de fragrância (`Eau de Parfum`, `EDP`, etc.)
2. Normaliza texto (lowercase, remove acentos, remove códigos numéricos)
3. Extrai palavras significativas (ignora stopwords como `de`, `da`, `ml`, `parfum`)
4. Calcula **Jaccard Score fuzzy** contra todos os produtos do CSV
5. Aceita se:
   - **Todas** as palavras do XLSX têm ao menos 1 par fuzzy no CSV, **E**
   - Score ≥ 0.35 **OU** cobertura 100%

### Passada 2 — Detecção de Conflitos
- Se 2+ produtos do XLSX apontam para o mesmo CSV → **conflito**
- Conflitos são **bloqueados** (não exportados) e exibidos no log com SKUs

---

## 🛠 Tecnologias

- **Python 3.8+**
- **Tkinter** — interface gráfica nativa
- **Pandas** — manipulação de dados CSV/XLSX
- **openpyxl** — leitura de arquivos Excel (.xlsx)
- **difflib** — similaridade de strings (SequenceMatcher)
- **threading + queue** — processamento assíncrono sem travar a UI

---

## 📝 Licença

Projeto interno de uso pessoal / comercial privado.
