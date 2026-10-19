# ==============================================================================
#  VERIFICAÇÃO AUTOMÁTICA DE DEPENDÊNCIAS
# ==============================================================================
#
#  Antes de tudo, garantimos que as bibliotecas externas (pandas e openpyxl)
#  estão instaladas. Isso é importante porque o usuário final pode rodar o
#  programa pela primeira vez sem ter feito `pip install` manualmente.
#
#  A ideia é: tenta importar → se falhar, instala automaticamente via pip.
#  Isso torna o programa "plug and play" — duplo clique e funciona.
# ==============================================================================
import sys
import subprocess

# Dicionário: chave = nome do módulo Python, valor = nome do pacote no pip.
# Na maioria dos casos são iguais, mas nem sempre (ex: Pillow se importa como PIL).
_DEPENDENCIAS = {"pandas": "pandas", "openpyxl": "openpyxl"}

def _verificar_e_instalar_dependencias():
    """
    Verifica se todas as dependências estão disponíveis.
    Se alguma faltar, tenta instalar via pip automaticamente.

    Nota para o júnior: usamos `importlib.import_module()` ao invés de
    `import` direto porque queremos testar a importação sem poluir o
    namespace global — só queremos saber se o módulo EXISTE, não usá-lo aqui.
    """
    import importlib
    faltando = []
    for modulo, pacote in _DEPENDENCIAS.items():
        try:
            importlib.import_module(modulo)
        except ImportError:
            faltando.append(pacote)
    if faltando:
        print(f"[AUTO-INSTALL] Instalando: {faltando}")
        try:
            # `sys.executable` garante que usamos o mesmo Python que está rodando
            # este script, evitando problemas quando há múltiplas versões instaladas.
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--quiet"] + faltando
            )
            print("[AUTO-INSTALL] Instalacao concluida!")
        except subprocess.CalledProcessError as e:
            print(f"[AUTO-INSTALL] FALHA: {e}")
            print(f"Execute manualmente: pip install {' '.join(faltando)}")
            sys.exit(1)

# Executa a verificação no momento da importação do módulo.
# Isso garante que quando a interface importar `logica_perfumaria`,
# as dependências já estarão prontas.
_verificar_e_instalar_dependencias()
# ==============================================================================

import re
import math
import difflib       # Para calcular similaridade entre strings (SequenceMatcher)
import unicodedata   # Para remover acentos de texto (normalização Unicode)
import os
import pandas as pd

# ==============================================================================
#  CONSTANTES GLOBAIS
# ==============================================================================

# Fator de conversão: 1 fl.oz. = 29.5735 ml
# Usado para gerar a descrição do fornecedor (ex: "PERFUME X 100ML 3.4 FL.OZ.")
ML_POR_FLOZ  = 29.5735

# Score mínimo do Jaccard para aceitar um match automático.
# 0.35 é relativamente baixo — mas combinamos com a regra de "cobertura total"
# para evitar falsos positivos. O Jaccard sozinho não basta porque produtos
# com 1-2 palavras distintas (ex: "Aabinus") ficam com score muito baixo.
SCORE_MINIMO = 0.35

# ==============================================================================
#  MAPEAMENTO FLEXÍVEL DE COLUNAS DO XLSX DO FORNECEDOR
# ==============================================================================
#
#  O problema: cada fornecedor nomeia as colunas de forma diferente.
#  Um pode usar "Nome do Perfume", outro "Produto", outro "Nome Formatado (ERP)".
#
#  A solução: para cada coluna que PRECISAMOS internamente, definimos uma lista
#  de "apelidos" aceitos. Na hora de ler o XLSX, procuramos por qualquer um
#  desses apelidos e mapeamos para o nome interno padronizado.
#
#  Nota: a comparação é case-insensitive e sem acentos (via `normalizar()`).
# ==============================================================================
MAPA_COLUNAS = {
    'Nome do Perfume': [
        'nome do perfume', 'nome formatado (erp)', 'produto',
        'nome', 'perfume', 'description',
    ],
    'Quantidade (ml)': [
        'quantidade (ml)', 'quantidade', 'ml', 'volume', 'qty',
    ],
    'Valor Unitario': [
        'valor unitario', 'valor unitário', 'valor unitario (us$)',
        'valor unitário (us$)', 'valor unit.', 'unit price', 'price', 'preco',
    ],
    # Coluna OPCIONAL — obrigatória apenas no Modo Estoque.
    # Se o usuário não marcou "Modo Estoque", a ausência desta coluna é ignorada.
    'Estoque (un)': [
        'estoque (un)', 'estoque', 'quantidade (un)', 'unidades',
        'qty estoque', 'stock', 'qtd estoque',
    ],
}

# ==============================================================================
#  PALAVRAS IGNORADAS NO MATCHING (stopwords do domínio de perfumaria)
# ==============================================================================
#
#  Essas palavras aparecem em quase todos os nomes de perfume e NÃO ajudam
#  a distinguir um produto do outro. Se não as removêssemos, dois perfumes
#  completamente diferentes como "Eau de Parfum X" e "Eau de Parfum Y"
#  teriam um Jaccard artificialmente alto por compartilharem 3 palavras.
#
#  CUIDADO: palavras como "intense" e "masc" NÃO estão aqui de propósito!
#  Elas foram removidas do conjunto _IGNORAR porque são DISTINTIVAS:
#  "Club Nuit Intense" é um produto diferente de "Club Nuit".
#  Essa foi uma lição aprendida na Fase 5 do desenvolvimento.
# ==============================================================================
_IGNORAR = {
    'de', 'da', 'do', 'the', 'eau', 'ml', 'oz', 'parfum', 'toilette',
    'perfume', 'edt', 'edp', 'e', 'a', 'o', 'em', 'com', 'perf',
    'fem', 'masc', 'feme', 'nuevo', 'new', 'branco', 'negro',
    'limited', 'edition', 'intense',
}

# Threshold para considerar duas palavras como "fuzzy match".
# 0.78 significa que 78% dos caracteres precisam coincidir.
# Exemplo: "Marshmallow" vs "Marshm" → prefixo, aceito diretamente.
# Exemplo: "Musaman" vs "Musamam" → ratio ~0.85, aceito pelo threshold.
_FUZZY_WORD_THRESHOLD = 0.78


# ==============================================================================
#  UTILITÁRIOS — Funções auxiliares usadas em todo o módulo
# ==============================================================================

def formatar_moeda(valor):
    """
    Converte float para string no formato brasileiro: "179,90"
    O Bling espera valores com vírgula como separador decimal.
    """
    return f"{valor:.2f}".replace('.', ',')


def parse_moeda(valor_str):
    """
    Faz o inverso do formatar_moeda: converte string de moeda para float.

    Lida com vários formatos que podem aparecer nos arquivos:
      - "R$ 179,90"  →  179.90
      - "1.234,56"   →  1234.56  (milhar com ponto, decimal com vírgula)
      - "15.00"      →  15.00    (formato americano)
      - "US$ 15,00"  →  15.00
      - ""           →  0.0      (célula vazia)

    Nota para o júnior: a ordem de verificação importa!
    Se tem AMBOS ponto e vírgula, o ponto é milhar e a vírgula é decimal.
    Se tem SÓ vírgula, ela é o separador decimal.
    """
    if pd.isna(valor_str) or str(valor_str).strip() in ('', 'nan'):
        return 0.0
    v = str(valor_str).strip()
    # Remove prefixos de moeda comuns
    for s in ['R$', '$', 'EUR', 'U$']:
        v = v.replace(s, '').strip()
    # Lógica de separadores:
    if '.' in v and ',' in v:
        # "1.234,56" → remove pontos (milhar), troca vírgula por ponto (decimal)
        v = v.replace('.', '').replace(',', '.')
    elif ',' in v:
        # "179,90" → troca vírgula por ponto
        v = v.replace(',', '.')
    try:
        return float(v)
    except ValueError:
        return 0.0


def arredondar_preco(valor):
    """
    Arredonda o preço para terminar sempre em ,90.

    Lógica: arredonda para a dezena mais próxima e subtrai 10 centavos.
    Exemplo: 181.50 → round(181.50/10)*10 = 180 → 180 - 0.10 = 179.90

    Isso é uma estratégia de precificação psicológica comum no varejo:
    R$ 179,90 "parece" mais barato que R$ 180,00.
    """
    return round(valor / 10) * 10 - 0.10


def converter_para_floz(ml):
    """Converte mililitros para fluid ounces (arredonda 1 casa)."""
    return round(ml / ML_POR_FLOZ, 1)


def gerar_descricao_fornecedor(nome_perfume, quantidade_ml):
    """
    Gera a string padronizada para o campo "Descrição do Produto no Fornecedor"
    no Bling. Formato: "NOME DO PERFUME 100ML 3.4 FL.OZ."

    Nota: se a quantidade vier como NaN (célula vazia no XLSX), assume 100ml
    como fallback defensivo. Isso evita um crash e na maioria dos casos
    perfumes são 100ml mesmo.
    """
    nome_upper = str(nome_perfume).upper().strip()
    try:
        ml = float(quantidade_ml)
        # Verifica NaN explicitamente — float('nan') passa pelo try/except
        # mas causa crash em operações posteriores como int(ml).
        if math.isnan(ml):
            ml = 100.0
    except (ValueError, TypeError):
        ml = 100.0
    floz = converter_para_floz(ml)
    # Se 100.0 → mostra "100ML", se 33.3 → mostra "33.3ML"
    ml_int = int(ml) if ml == int(ml) else ml
    return f"{nome_upper} {ml_int}ML {floz} FL.OZ."


def normalizar(texto):
    """
    Normaliza texto para comparação: lowercase + remove acentos.

    Usa decomposição Unicode NFKD: "é" vira "e" + acento combinante,
    depois filtra só os caracteres ASCII (remove os acentos).

    Isso é essencial para o matching porque "Descrição" == "descricao"
    quando normalizado.
    """
    s = str(texto).lower().strip()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return s


def strip_codigos(texto):
    """
    Remove tokens que começam com dígito (códigos como "2230", "100M6173").

    Esses códigos são internos do fornecedor e não existem no CSV do Bling,
    então se ficassem no matching, seriam "palavras sem par" e bloqueariam
    matches válidos.

    Exemplo: "2230 Asad Lattafa 100ml" → "Asad Lattafa 100ml"
    """
    return ' '.join(t for t in texto.split() if not re.match(r'^\d', t))


# ==============================================================================
#  REMOÇÃO DE PREFIXOS DE FRAGRÂNCIA
# ==============================================================================
#
#  Nomes de perfume frequentemente começam com o tipo de fragrância:
#    "Eau de Parfum Aseel Maison Asrar 100ml"
#    "PERF Aseel Maison Asrar 100ml"
#    "EDP Me L Affair 100ml"
#
#  O problema: se não removermos esses prefixos, as palavras "eau", "de",
#  "parfum" entram no cálculo do Jaccard e inflam o score sem agregar
#  valor na distinção entre produtos.
#
#  A regex abaixo captura TODAS as variações comuns no início do texto
#  e remove apenas a PRIMEIRA ocorrência (count=1).
# ==============================================================================
_PREFIXOS_FRAGANCIA = re.compile(
    r'^(eau\s+de\s+(parfum|toilette|cologne)|extrait\s+de\s+parfum|'
    r'eau\s+de\s+parfum|edp|edt|edc|parfum|perfume|perf|cologne)\s+',
    re.IGNORECASE
)


def _strip_prefixo(texto):
    """
    Remove prefixo de tipo de fragrância do início do texto.
    Exemplos:
      'Eau de Parfum Aseel Maison Asrar 100ml' -> 'Aseel Maison Asrar 100ml'
      'PERF Aseel Maison Asrar 100ml'          -> 'Aseel Maison Asrar 100ml'
      'EDP Me L Affair 100ml'                  -> 'Me L Affair 100ml'
    """
    return _PREFIXOS_FRAGANCIA.sub('', texto.strip(), count=1).strip()


def palavras_sig(texto):
    """
    Extrai o conjunto de "palavras significativas" de um nome de perfume.

    Pipeline:
      1. Remove prefixo de fragrância ("Eau de Parfum ...")
      2. Normaliza (lowercase, sem acentos)
      3. Remove códigos numéricos
      4. Filtra stopwords e palavras muito curtas (≤ 2 chars)

    O resultado é um SET — a ordem não importa, o que é crucial porque
    o CSV pode ter "Eau de Parfum Asad Lattafa 100ml" e o XLSX pode ter
    "Asad Lattafa" (sem as mesmas palavras, sem a mesma ordem).

    Retorna set para uso direto no cálculo Jaccard.
    """
    limpo = strip_codigos(normalizar(_strip_prefixo(texto)))
    return {p for p in limpo.split() if p not in _IGNORAR and len(p) > 2}


def _palavras_fuzzy_match(pw_a, pw_b):
    """
    Verifica se duas palavras são "fuzzy match" — ou seja, similares o bastante
    para serem consideradas a mesma palavra.

    Duas estratégias:
      1. Prefixo: "Marshm" é prefixo de "Marshmallow" → match direto
      2. Similaridade: usa SequenceMatcher do difflib para calcular ratio.
         "Musaman" vs "Musamam" → ratio ~0.85 → aceito (threshold = 0.78)

    Por que prefixo? Fornecedores frequentemente abreviam nomes longos.
    """
    if pw_b.startswith(pw_a) or pw_a.startswith(pw_b):
        return True
    return difflib.SequenceMatcher(None, pw_a, pw_b).ratio() >= _FUZZY_WORD_THRESHOLD


# ==============================================================================
#  SCORING Jaccard fuzzy-aware  (0.0 a 1.0)
# ==============================================================================
#
#  O Jaccard clássico é: |A ∩ B| / |A ∪ B|
#  Mas aqui usamos uma versão "fuzzy": duas palavras que são similares
#  (via _palavras_fuzzy_match) contam como interseção.
#
#  Exemplo:
#    XLSX: {"asad", "lattafa"}        (2 palavras)
#    CSV:  {"asad", "lattafa", "100ml"} (3 palavras, mas "100ml" é filtrada)
#    → casadas = 2, união = 2 + 0 = 2, score = 1.0
#
#  Se o CSV tivesse palavras extras que NÃO casam:
#    CSV: {"asad", "lattafa", "gold"}
#    → casadas = 2, csv_não_casadas = {"gold"}, união = 2 + 1 = 3, score = 0.67
# ==============================================================================

def _jaccard_score(p_xlsx: set, p_csv: set) -> float:
    """
    Calcula score Jaccard fuzzy entre dois conjuntos de palavras.

    Retorna um float de 0.0 (nenhuma similaridade) a 1.0 (match perfeito).
    A comparação entre palavras individuais usa fuzzy matching (prefixo ou
    similaridade de 78%+).
    """
    if not p_xlsx or not p_csv:
        return 0.0

    casadas_xlsx = 0       # Quantas palavras do XLSX encontraram par no CSV
    casadas_csv  = set()   # Quais palavras do CSV foram "consumidas" por um match

    for pw in p_xlsx:
        for pc in p_csv:
            if _palavras_fuzzy_match(pw, pc):
                casadas_xlsx += 1
                casadas_csv.add(pc)
                break   # Cada palavra do XLSX casa com no máximo 1 do CSV

    # União = palavras do XLSX + palavras do CSV que NÃO foram casadas
    # Isso penaliza CSVs com muitas palavras extras (nomes mais genéricos)
    uniao = len(p_xlsx) + len(p_csv - casadas_csv)
    return casadas_xlsx / uniao if uniao > 0 else 0.0


# ==============================================================================
#  LEITURA DO XLSX DO FORNECEDOR
# ==============================================================================

def _detectar_colunas(df_colunas_reais, log_callback, modo_estoque=False):
    """
    Detecta quais colunas do XLSX correspondem às colunas internas esperadas
    usando o MAPA_COLUNAS de apelidos.

    Retorna um dicionário {nome_interno: nome_real_no_xlsx} ou None se
    alguma coluna obrigatória não for encontrada.

    Nota: no modo estoque, "Estoque (un)" também é obrigatória.
    """
    mapa = {}
    # Cria um índice normalizado das colunas reais para busca case-insensitive
    colunas_lower = {normalizar(c): c for c in df_colunas_reais}
    obrigatorias  = ['Nome do Perfume', 'Quantidade (ml)', 'Valor Unitario']
    if modo_estoque:
        obrigatorias.append('Estoque (un)')

    for nome_interno, apelidos in MAPA_COLUNAS.items():
        encontrado = None
        for apelido in apelidos:
            if normalizar(apelido) in colunas_lower:
                encontrado = colunas_lower[normalizar(apelido)]
                break
        if encontrado:
            mapa[nome_interno] = encontrado
            # Loga quando o nome real difere do esperado, para o usuário saber
            # qual mapeamento foi feito (transparência)
            if normalizar(encontrado) != normalizar(nome_interno):
                log_callback(f"INFO: '{nome_interno}' mapeada de '{encontrado}'.", "alerta")
        elif nome_interno in obrigatorias:
            # Se a coluna é obrigatória e não foi encontrada, é erro fatal
            log_callback(f"[ERRO] Coluna obrigatória '{nome_interno}' não encontrada no XLSX.", "erro")
            log_callback(f"       Disponíveis: {list(df_colunas_reais)}", "erro")
            log_callback(f"       Aceitos: {MAPA_COLUNAS[nome_interno]}", "erro")
            return None
    return mapa


def ler_xlsx_fornecedor(caminho_xlsx, log_callback, modo_estoque=False):
    """
    Lê o XLSX do fornecedor, detecta e renomeia colunas, remove linhas
    sem nome de perfume.

    Retorna DataFrame padronizado ou None em caso de erro.

    Nota para o júnior: usamos `dtype=str` na leitura para evitar que o
    Pandas converta automaticamente valores numéricos. Isso previne
    problemas como "2230" (código) virar 2230.0 (float) e causar erros
    de matching ou formatação.
    """
    if not os.path.exists(caminho_xlsx):
        log_callback(f"[ERRO] XLSX não encontrado: {caminho_xlsx}", "erro")
        return None
    try:
        df = pd.read_excel(caminho_xlsx, dtype=str)
    except Exception as e:
        log_callback(f"[ERRO] Falha ao ler XLSX: {e}", "erro")
        return None

    mapa = _detectar_colunas(df.columns, log_callback, modo_estoque)
    if mapa is None:
        return None

    # Renomeia colunas do XLSX para os nomes internos padronizados.
    # Só renomeia se o nome real for diferente do interno (evita erro).
    renomear = {v: k for k, v in mapa.items() if normalizar(v) != normalizar(k)}
    if renomear:
        df = df.rename(columns=renomear)

    # Tratamento especial: "Valor Unitário" (com acento) → "Valor Unitario" (sem).
    # Isso pode acontecer quando o XLSX usa o nome interno com acento.
    for col in list(df.columns):
        if normalizar(col) == normalizar('Valor Unitário') and col != 'Valor Unitario':
            df = df.rename(columns={col: 'Valor Unitario'})
            break

    # Remove linhas onde o nome do perfume está vazio (lixo na planilha)
    df = df.dropna(subset=['Nome do Perfume'])
    log_callback(f"INFO: {len(df)} perfumes encontrados na planilha do fornecedor.", "sucesso")
    return df


# ==============================================================================
#  PASSADA 1 — Mapeamento Jaccard
# ==============================================================================
#
#  Para cada perfume no XLSX do fornecedor, procura o "melhor match" no CSV
#  do Bling usando Jaccard fuzzy scoring.
#
#  Critério de aceitação (definido após várias iterações de debug):
#    1. TODAS as palavras do XLSX precisam ter ao menos 1 par fuzzy no CSV.
#       → Isso evita "Aseel Maison Asrar" bater em "Rey Maison Asrar"
#         apenas porque "Maison" + "Asrar" casam (mas "Aseel" não tem par).
#    2. Score ≥ SCORE_MINIMO (0.35) OU cobertura 100%.
#       → A regra de cobertura resolve o caso do "Aabinus": 1 palavra única
#         com Jaccard 0.33 (abaixo do mínimo), mas com 100% de cobertura
#         (a única palavra do XLSX foi encontrada no CSV).
# ==============================================================================

def _passada_mapeamento(df_forn, csv_palavras, log_callback):
    """
    Primeira passada do matching: para cada produto do XLSX, encontra o
    melhor candidato no CSV via Jaccard scoring.

    Parâmetros:
      df_forn       — DataFrame do fornecedor (XLSX)
      csv_palavras  — dict {csv_idx: set_de_palavras_significativas}
      log_callback  — função para enviar mensagens ao log

    Retorna:
      xlsx_mapa       — {i_xlsx: (csv_idx, score, nome, qtd_ml, val_unit, qtd_estoque)}
      nao_encontrados — lista de nomes que não passaram nos critérios
    """
    xlsx_mapa       = {}
    nao_encontrados = []
    col_val     = 'Valor Unitario' if 'Valor Unitario' in df_forn.columns else 'Valor Unitário'
    tem_estoque = 'Estoque (un)' in df_forn.columns

    for i_xlsx, linha in df_forn.iterrows():
        nome_perfume       = str(linha['Nome do Perfume']).strip()
        quantidade_ml_str  = str(linha['Quantidade (ml)']).strip()
        valor_unitario_str = str(linha.get(col_val, '0')).strip()
        est_raw            = str(linha.get('Estoque (un)', '')).strip() if tem_estoque else ''

        # ── Validação de dados numéricos ──────────────────────────────
        # Células vazias no XLSX viram NaN no Pandas. Um NaN passa pelo
        # try/except do float() mas causa crash em operações posteriores
        # como int(ml). Por isso verificamos explicitamente com math.isnan().
        try:
            quantidade_ml = float(quantidade_ml_str.replace(',', '.'))
            if math.isnan(quantidade_ml):
                raise ValueError("NaN")
        except (ValueError, TypeError):
            log_callback(f"[AVISO] Quantidade inválida para '{nome_perfume}'. Pulando.", "erro")
            continue

        valor_unitario = parse_moeda(valor_unitario_str)
        if math.isnan(valor_unitario):
            log_callback(f"[AVISO] Valor unitário inválido para '{nome_perfume}'. Pulando.", "erro")
            continue
        qtd_estoque = parse_moeda(est_raw) if (tem_estoque and est_raw.lower() not in ('', 'nan', 'none')) else None

        # ── Busca do melhor match no CSV ──────────────────────────────
        p_xlsx       = palavras_sig(nome_perfume)
        melhor_idx   = None
        melhor_score = 0.0

        # Compara contra TODOS os produtos do CSV (força bruta).
        # Com ~500 produtos é instantâneo. Se fosse milhares,
        # poderíamos usar índice invertido, mas aqui não precisa.
        for csv_idx, p_csv in csv_palavras.items():
            score = _jaccard_score(p_xlsx, p_csv)
            if score > melhor_score:
                melhor_score = score
                melhor_idx   = csv_idx

        # ── Verifica critérios de aceitação ───────────────────────────
        cobertura_xlsx  = 0.0
        todas_casadas   = False   # toda palavra do XLSX tem ao menos 1 par no CSV
        if melhor_idx is not None and p_xlsx:
            p_csv_melhor = csv_palavras[melhor_idx]
            casadas_cob  = 0
            palavras_sem_par = []
            for pw in p_xlsx:
                if any(_palavras_fuzzy_match(pw, pc) for pc in p_csv_melhor):
                    casadas_cob += 1
                else:
                    palavras_sem_par.append(pw)
            cobertura_xlsx = casadas_cob / len(p_xlsx)
            todas_casadas  = (len(palavras_sem_par) == 0)
        else:
            palavras_sem_par = []

        # Critério de aceitação (ver comentário do bloco acima para detalhes):
        #   1. Toda palavra do XLSX precisa ter ao menos 1 par fuzzy no CSV.
        #      → Evita "Aseel" bater em "Rey Maison Asrar" apenas via "Maison"+"Asrar".
        #   2. Score >= mínimo  OU  cobertura 100% (regra do "Aabinus" — 1 palavra única).
        aceitar = (
            melhor_idx is not None
            and todas_casadas
            and (melhor_score >= SCORE_MINIMO or cobertura_xlsx == 1.0)
        )

        if not aceitar:
            nao_encontrados.append(nome_perfume)
            motivo = (
                f"palavra(s) sem par no CSV: {palavras_sem_par}"
                if not todas_casadas
                else f"score: {melhor_score:.2f}, cobertura: {cobertura_xlsx:.0%}"
            )
            log_callback(
                f"NÃO ENCONTRADO: '{nome_perfume}' ({motivo})",
                "erro"
            )
            continue

        xlsx_mapa[i_xlsx] = (melhor_idx, melhor_score, nome_perfume,
                             quantidade_ml, valor_unitario, qtd_estoque)
    return xlsx_mapa, nao_encontrados


# ==============================================================================
#  FUNÇÃO PRINCIPAL
#  Retorna resultado (dict) sem salvar — a interface salva após o preview.
# ==============================================================================
#
#  Essa é a função "orquestradora". Ela:
#    1. Lê o CSV do Bling
#    2. Lê o XLSX do fornecedor
#    3. Roda a Passada 1 (mapeamento Jaccard)
#    4. Detecta conflitos (Passada 2)
#    5. Aplica preços ou estoque em memória (sem salvar em disco!)
#    6. Monta os dados para o Preview
#    7. Retorna tudo num dicionário para a interface tratar
#
#  IMPORTANTE: ela NÃO grava o CSV final. Isso é feito pela `salvar_resultado()`
#  somente APÓS o usuário confirmar no Preview. Isso dá ao usuário a chance de
#  revisar, trocar matches errados e remover linhas antes de exportar.
# ==============================================================================

def processar_perfumaria(arquivo_csv, arquivo_xlsx, cotacao_dolar, taxa_frete,
                         log_callback, modo_estoque=False, limite_variacao=10.0):
    """
    Processa e retorna dict com resultado (sem salvar em disco).
    Retorna None em caso de erro fatal.

    Chaves do resultado:
      df_bling, indices_alterados, caminho_saida, preview_rows,
      conflitos_idx, nao_encontrados, ok_count, modo
    """
    # --- Leitura do CSV do Bling ---
    if not os.path.exists(arquivo_csv):
        log_callback("[ERRO] CSV do Bling não encontrado.", "erro")
        return None
    try:
        # O CSV do Bling usa separador ; (ponto-e-vírgula) e encoding utf-8 com BOM.
        # utf-8-sig lê e ignora o BOM automaticamente.
        df_bling = pd.read_csv(arquivo_csv, sep=';', encoding='utf-8-sig', dtype=str)
    except Exception as e:
        log_callback(f"[ERRO] Falha ao ler CSV: {e}", "erro")
        return None

    # O Bling pode exportar a coluna como "Descrição" ou "Descricao" (sem acento).
    # Procuramos a que existir.
    col_desc = next((c for c in ('Descrição', 'Descricao') if c in df_bling.columns), None)
    if col_desc is None:
        log_callback("[ERRO] Coluna 'Descrição' não encontrada no CSV.", "erro")
        return None
    # Preenche NaN na descrição para evitar crashes em operações de string
    df_bling[col_desc] = df_bling[col_desc].fillna('Produto sem nome')

    # --- Leitura do XLSX ---
    df_forn = ler_xlsx_fornecedor(arquivo_xlsx, log_callback, modo_estoque)
    if df_forn is None:
        return None

    # --- Helpers para acessar SKU e GTIN de forma segura ---
    col_sku  = next((c for c in ('Código', 'Codigo') if c in df_bling.columns), None)
    col_gtin = 'GTIN/EAN' if 'GTIN/EAN' in df_bling.columns else None

    def get_sku(idx):
        """Retorna o SKU da linha idx do CSV, ou 'N/A' se a coluna não existir."""
        return str(df_bling.at[idx, col_sku]).strip() if col_sku else 'N/A'

    def get_gtin(idx):
        """
        Retorna o GTIN/EAN da linha idx, ou '' se não tiver.
        Filtra valores inválidos como 'nan', 'Sem GTIN', 'Sem EAN'.
        """
        if col_gtin is None:
            return ''
        val = str(df_bling.at[idx, col_gtin]).strip()
        return '' if val.lower() in ('nan', '', 'sem gtin', 'sem ean') else val

    # Pré-calcula as palavras significativas de TODOS os produtos do CSV.
    # Isso é feito UMA VEZ e reutilizado em todas as comparações Jaccard.
    csv_palavras = {idx: palavras_sig(row[col_desc]) for idx, row in df_bling.iterrows()}

    # ------------------------------------------------------------------
    #  PASSADA 1 — Encontra o melhor match para cada XLSX
    # ------------------------------------------------------------------
    log_callback("\n--- PASSADA 1: MAPEAMENTO (Jaccard Scoring) ---")
    xlsx_mapa, nao_encontrados = _passada_mapeamento(df_forn, csv_palavras, log_callback)

    # ------------------------------------------------------------------
    #  DETECÇÃO DE CONFLITOS
    # ------------------------------------------------------------------
    #  Verifica se 2+ produtos do XLSX apontam para o MESMO produto do CSV.
    #  Exemplo: "DOLCE G LIGHT BLUE FEME EDT" e "DOLCE G FEME EDT" ambos
    #  apontam para o SKU "ABC123" → conflito, nenhum dos dois é exportado.
    #
    #  Isso evita sobrescrever silenciosamente: sem essa verificação,
    #  o último XLSX processado ganharia e o anterior seria perdido.
    # ------------------------------------------------------------------
    csv_reivindicado = {}
    for i_xlsx, (csv_idx, *_) in xlsx_mapa.items():
        csv_reivindicado.setdefault(csv_idx, []).append(i_xlsx)
    indices_conflito = {csv_idx for csv_idx, rv in csv_reivindicado.items() if len(rv) > 1}

    # Move os produtos do XLSX que entraram em conflito para os "não encontrados"
    # Assim eles aparecem na tela de reassociação manual para o usuário resolver.
    for csv_idx in indices_conflito:
        for i_xlsx in csv_reivindicado[csv_idx]:
            nome_perf_conflito = xlsx_mapa[i_xlsx][2]
            if nome_perf_conflito not in nao_encontrados:
                nao_encontrados.append(nome_perf_conflito)

    # ------------------------------------------------------------------
    #  PASSADA 2 — Aplica em memória (preço ou estoque)
    # ------------------------------------------------------------------
    log_callback(f"\n--- PASSADA 2: {'MODO ESTOQUE' if modo_estoque else 'APLICAÇÃO DE PREÇOS'} ---")

    indices_alterados = set()    # Índices do CSV que foram modificados (para exportar depois)
    preview_rows      = []       # Dados para a tabela de preview na interface
    ok_count          = 0        # Contador de produtos processados com sucesso

    for i_xlsx, (csv_idx, score, nome_perfume, quantidade_ml,
                 valor_unitario, qtd_estoque) in xlsx_mapa.items():

        desc_csv = df_bling.at[csv_idx, col_desc]
        sku_csv  = get_sku(csv_idx)
        gtin_csv = get_gtin(csv_idx)

        # Pula produtos em conflito — serão reportados no final
        if csv_idx in indices_conflito:
            continue

        # ── MODO ESTOQUE ──────────────────────────────────────────────
        if modo_estoque:
            if qtd_estoque is None:
                log_callback(
                    f"[AVISO] Sem coluna 'Estoque (un)' no XLSX para '{nome_perfume}'. Pulando.",
                    "erro"
                )
                continue

            # Lê o estoque atual do CSV e SOMA com a nova quantidade.
            estoque_antigo_str = str(df_bling.at[csv_idx, 'Estoque']).strip() \
                if 'Estoque' in df_bling.columns else '0'
            estoque_antigo = parse_moeda(estoque_antigo_str) or 0.0
            estoque_novo = estoque_antigo + qtd_estoque
            qtd_fmt = str(int(estoque_novo)) if estoque_novo == int(estoque_novo) \
                else f"{estoque_novo:.2f}".replace('.', ',')
            estoque_antigo = estoque_antigo_str  # mantém string original para o log

            if 'Estoque' in df_bling.columns:
                df_bling.at[csv_idx, 'Estoque'] = qtd_fmt

            msg_gtin = f" ⚠️ GTIN: {gtin_csv}" if gtin_csv else ""
            log_callback(
                f"ESTOQUE [SKU {sku_csv}]: {desc_csv} | "
                f"{estoque_antigo} → {qtd_fmt} un{msg_gtin}",
                "gtin" if gtin_csv else "sucesso"
            )
            preview_rows.append({
                'sku'      : sku_csv,
                'csv_idx'  : csv_idx,
                'descricao': desc_csv,
                'campo'    : 'Estoque',
                'antes'    : estoque_antigo,
                'depois'   : qtd_fmt,
                'gtin'     : gtin_csv,
                'qtd_forn' : qtd_estoque,
            })

        # ── MODO PREÇO ────────────────────────────────────────────────
        else:
            # Cálculo do preço:
            #   Compra = valor_unitario × cotação_dolar
            #   Custo  = compra × (1 + frete%)
            #   Venda  = arredonda(custo × 2, terminando em ,90)
            preco_compra      = valor_unitario * cotacao_dolar
            preco_custo       = preco_compra * (1 + taxa_frete / 100.0)
            preco_venda_final = arredondar_preco(preco_custo * 2)
            desc_fornecedor   = gerar_descricao_fornecedor(nome_perfume, quantidade_ml)

            preco_antigo = str(df_bling.at[csv_idx, 'Preço']).strip() \
                if 'Preço' in df_bling.columns else '0,00'
            preco_novo   = formatar_moeda(preco_venda_final)

            # Atualiza as colunas de preço no DataFrame (em memória, NÃO no disco)
            if 'Preço de Compra' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço de Compra'] = formatar_moeda(preco_compra)
            if 'Preço de custo' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço de custo']  = formatar_moeda(preco_custo)
            if 'Preço' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço']            = preco_novo
            if 'Descrição do Produto no Fornecedor' in df_bling.columns:
                df_bling.at[csv_idx, 'Descrição do Produto no Fornecedor'] = desc_fornecedor

            # ── Diferença de preço (usado para cor no preview e log) ───
            diff_preco = abs(preco_venda_final - parse_moeda(preco_antigo))

            # ── Escolhe tag de cor para o log ─────────────────────────
            #  A lógica de cores só se aplica a produtos com GTIN/EAN.
            #  Produtos com GTIN são listados em comparadores de preço (Google Shopping,
            #  Buscapé, etc.), então variações grandes de preço são mais sensíveis.
            #
            #  Sem GTIN → sempre verde (sucesso normal).
            #  Com GTIN:
            #    diff == 0          → branco (sem destaque, nada mudou)
            #    0 < diff ≤ 10     → amarelo (variação pequena, provavelmente ok)
            #    10 < diff < 80    → laranja (atenção, revisar)
            #    diff >= 80        → vermelho (variação grande, perigo!)
            if not gtin_csv:
                tag_log = "sucesso"
            elif diff_preco == 0:
                tag_log = "sucesso"
            elif diff_preco <= 10:
                tag_log = "gtin_amarelo"
            elif diff_preco < 80:
                tag_log = "gtin_laranja"
            else:
                tag_log = "gtin_vermelho"

            if gtin_csv and diff_preco > 0:
                log_callback(
                    f"⚠️  GTIN [{gtin_csv}] | SKU {sku_csv} | {desc_csv} "
                    f"| XLSX: '{nome_perfume}' "
                    f"| R$ {preco_antigo} → R$ {preco_novo} (Δ R$ {diff_preco:.2f})",
                    tag_log
                )
            else:
                log_callback(
                    f"OK [SKU {sku_csv}]: {desc_csv} | "
                    f"R$ {preco_antigo} -> R$ {preco_novo} | "
                    f"[{desc_fornecedor}] (score: {score:.2f})",
                    tag_log
                )

            # Monta os dados para uma linha do Preview
            preview_rows.append({
                'sku'              : sku_csv,
                'csv_idx'          : csv_idx,
                'descricao'        : desc_csv,
                'campo'            : 'Preço',
                'antes'            : preco_antigo,
                'depois'           : preco_novo,
                'gtin'             : gtin_csv,
                'diff_preco'       : diff_preco,
                'variacao_pequena' : (0 < diff_preco <= limite_variacao),
                'custo'            : formatar_moeda(preco_custo),
                'compra'           : formatar_moeda(preco_compra),
                'desc_forn'        : desc_fornecedor,
                'nome_xlsx'        : nome_perfume,
                'quantidade_ml'    : quantidade_ml,
                'valor_unitario'   : valor_unitario,
                'qtd_forn'         : qtd_estoque,
            })

        indices_alterados.add(csv_idx)
        ok_count += 1

    # --- Relatório final de conflitos e não encontrados ---
    log_callback("\n" + "=" * 55)
    if indices_conflito:
        log_callback(
            f"\n[CONFLITO] {len(indices_conflito)} produto(s) do CSV disputado(s) — "
            f"os itens do XLSX foram enviados para reassociação manual:",
            "alerta"
        )
        for csv_idx in sorted(indices_conflito):
            sku_csv    = get_sku(csv_idx)
            desc_csv   = df_bling.at[csv_idx, col_desc]
            candidatos = csv_reivindicado[csv_idx]
            log_callback(f"\n  X SKU {sku_csv} | '{desc_csv}' — disputado por:", "erro")
            for i_xlsx in candidatos:
                log_callback(
                    f"      * '{xlsx_mapa[i_xlsx][2]}' (score: {xlsx_mapa[i_xlsx][1]:.2f})",
                    "erro"
                )

    if nao_encontrados:
        log_callback(f"\nNÃO ENCONTRADOS: {len(nao_encontrados)} perfume(s):", "erro")
        for nome in nao_encontrados:
            log_callback(f"  -> '{nome}'", "erro")
    else:
        log_callback("SUCESSO: Todos os itens do XLSX foram localizados no CSV!", "sucesso")

    log_callback(f"\nRESUMO:", "alerta")
    log_callback(f"  Prontos p/ exportar  : {ok_count}", "sucesso")
    log_callback(f"  Bloqueados (conflito): {len(indices_conflito)}",
                 "alerta" if indices_conflito else None)
    log_callback(f"  Não encontrados      : {len(nao_encontrados)}",
                 "erro" if nao_encontrados else None)
    log_callback("  ⬆ Verifique o Preview e confirme para salvar.", "info")
    log_callback("=" * 55)

    # ------------------------------------------------------------------
    #  Prepara dados dos "não encontrados" para reassociação manual na UI
    # ------------------------------------------------------------------
    #  Para cada item que não passou no matching automático, empacotamos
    #  seus dados (nome, ml, valor, estoque) num dicionário para que a
    #  janela ReassociarWindow possa mostrá-los ao usuário.
    # ------------------------------------------------------------------
    nao_encontrados_data = []
    col_val_forn = next((c for c in ('Valor Unitario', 'Valor Unitário')
                         if c in df_forn.columns), None)
    tem_estoque_forn = 'Estoque (un)' in df_forn.columns
    for i_xlsx, linha in df_forn.iterrows():
        nome = str(linha['Nome do Perfume']).strip()
        if nome in nao_encontrados:
            est_raw = str(linha.get('Estoque (un)', '')).strip() if tem_estoque_forn else ''
            qtd_est = parse_moeda(est_raw) if (tem_estoque_forn and est_raw.lower() not in ('', 'nan', 'none')) else None
            nao_encontrados_data.append({
                'i_xlsx'        : i_xlsx,
                'nome_perfume'  : nome,
                'quantidade_ml' : float(str(linha.get('Quantidade (ml)', '100')).replace(',', '.')),
                'valor_unitario': parse_moeda(str(linha.get(col_val_forn or '', '0'))),
                'qtd_estoque'   : qtd_est,
            })

    # Lista de todas as descrições do CSV para popular os comboboxes
    # de reassociação manual e troca de match
    csv_descricoes = [(idx, str(row[col_desc]).strip()) for idx, row in df_bling.iterrows()]

    caminho_saida = os.path.join(os.path.dirname(arquivo_csv), 'perfumes_atualizados.csv')
    caminho_txt   = os.path.join(os.path.dirname(arquivo_csv), 'estoque.txt')
    return {
        'df_bling'            : df_bling,
        'indices_alterados'   : indices_alterados,
        'caminho_saida'       : caminho_saida,
        'caminho_txt'         : caminho_txt,
        'tem_estoque_forn'    : tem_estoque_forn,
        'preview_rows'        : preview_rows,
        'conflitos_idx'       : indices_conflito,
        'nao_encontrados'     : nao_encontrados,
        'nao_encontrados_data': nao_encontrados_data,
        'csv_descricoes'      : csv_descricoes,
        'ok_count'            : ok_count,
        'modo'                : 'estoque' if modo_estoque else 'preco',
        'cotacao_dolar'       : cotacao_dolar,
        'taxa_frete'          : taxa_frete,
        'limite_variacao'     : limite_variacao,
        'col_desc'            : col_desc,
        'col_sku'             : col_sku,
        'col_gtin'            : col_gtin,
    }


def salvar_resultado(resultado, log_callback):
    """
    Grava o CSV final após confirmação no preview.

    Exporta SOMENTE as linhas que foram alteradas (indices_alterados),
    não o CSV inteiro. O arquivo é salvo na mesma pasta do CSV original
    com o nome 'perfumes_atualizados.csv'.

    Se o Modo Estoque estiver ativo ou a planilha do fornecedor possuir
    coluna de estoque, gera também 'estoque.txt' no formato:
      SKU: {sku} | Estoque: {quantidade}
    """
    df_bling          = resultado['df_bling']
    indices_alterados = resultado['indices_alterados']
    caminho_saida     = resultado['caminho_saida']
    caminho_txt       = resultado.get('caminho_txt') or os.path.join(
        os.path.dirname(caminho_saida), 'estoque.txt'
    )
    modo_estoque      = resultado.get('modo') == 'estoque'
    tem_estoque_forn  = resultado.get('tem_estoque_forn', False)

    if not indices_alterados:
        log_callback("AVISO: Nenhum produto para exportar.", "erro")
        return False

    df_saida = df_bling.loc[list(indices_alterados)].copy()
    df_saida.to_csv(caminho_saida, sep=';', encoding='utf-8-sig', index=False)
    log_callback(f"✅ ARQUIVO SALVO: {caminho_saida}", "sucesso")
    log_callback(f"   {len(indices_alterados)} linha(s) exportadas.", "sucesso")

    # Exportação do estoque.txt se estiver no Modo Estoque ou se o XLSX tiver coluna de estoque
    if modo_estoque or tem_estoque_forn:
        preview_rows = resultado.get('preview_rows', [])
        estoque_por_sku = {}
        for r in preview_rows:
            if r.get('sel', '☑') == '☐':
                continue
            sku = str(r.get('sku', '')).strip()
            qtd = r.get('qtd_forn')
            if not sku or sku == 'N/A' or qtd is None:
                continue
            try:
                qtd_num = float(qtd)
            except (ValueError, TypeError):
                continue
            estoque_por_sku[sku] = estoque_por_sku.get(sku, 0.0) + qtd_num

        if estoque_por_sku:
            try:
                with open(caminho_txt, 'w', encoding='utf-8') as f:
                    for sku, qtd in estoque_por_sku.items():
                        qtd_str = str(int(qtd)) if int(qtd) == qtd else f"{qtd:.2f}".replace('.', ',')
                        f.write(f"SKU: {sku} | Estoque: {qtd_str}\n")
                    f.write("Nota ;\n")
                    f.write("Entrada de Produtos ;\n")
                    f.write("Nome: Julio;\n")
                    f.write("QTDE.:02x\n")
                log_callback(f"✅ ESTOQUE TXT SALVO: {caminho_txt}", "sucesso")
                log_callback(f"   {len(estoque_por_sku)} produto(s) no estoque TXT.", "sucesso")
            except Exception as e:
                log_callback(f"[ERRO] Falha ao salvar estoque.txt: {e}", "erro")
        else:
            log_callback("ℹ️ Nenhum produto com estoque informado no XLSX para o estoque.txt.", "alerta")

    return True


def aplicar_reassociacao(resultado, reassociacoes, log_callback):
    """
    Aplica associações manuais (da ReassociarWindow ou TrocarMatchDialog)
    ao resultado, atualizando df_bling, preview_rows e indices_alterados.

    Esta função é chamada quando o usuário associa manualmente um produto
    do XLSX a um produto do CSV que o matching automático não encontrou.

    Os cálculos de preço são idênticos aos da Passada 2 — replicamos a
    fórmula aqui para manter consistência:
      Compra = valor_unitario × cotação_dolar
      Custo  = compra × (1 + frete%)
      Venda  = arredonda(custo × 2)
    """
    df_bling        = resultado['df_bling']
    cotacao_dolar   = resultado.get('cotacao_dolar', 1.0)
    taxa_frete      = resultado.get('taxa_frete', 0.0)
    limite_variacao = resultado.get('limite_variacao', 10.0)
    modo_estoque    = resultado.get('modo') == 'estoque'
    col_desc        = resultado.get('col_desc', 'Descrição')
    col_sku         = resultado.get('col_sku')
    col_gtin        = resultado.get('col_gtin')

    def get_sku(idx):
        return str(df_bling.at[idx, col_sku]).strip() if col_sku else 'N/A'

    def get_gtin(idx):
        if not col_gtin:
            return ''
        val = str(df_bling.at[idx, col_gtin]).strip()
        return '' if val.lower() in ('nan', '', 'sem gtin', 'sem ean') else val

    def fmt_qtd(v):
        """Formata quantidade: inteiro se possível, senão 2 casas com vírgula."""
        return str(int(v)) if v == int(v) else f"{v:.2f}".replace('.', ',')

    for r in reassociacoes:
        csv_idx        = r['csv_idx']
        nome_perfume   = r['nome_perfume']
        quantidade_ml  = r.get('quantidade_ml', 100.0)
        valor_unitario = r.get('valor_unitario', 0.0)
        qtd_estoque    = r.get('qtd_estoque')

        # Se o produto foi reassociado com sucesso, removemos da lista de não encontrados
        if 'nao_encontrados' in resultado and nome_perfume in resultado['nao_encontrados']:
            resultado['nao_encontrados'].remove(nome_perfume)

        desc_csv = str(df_bling.at[csv_idx, col_desc]).strip()
        sku_csv  = get_sku(csv_idx)
        gtin_csv = get_gtin(csv_idx)

        if modo_estoque:
            # No modo estoque da reassociação, SOMAMOS ao estoque existente
            # (diferente da passada 2 onde substituímos).
            # Isso porque reassociação manual é para "adicionar" estoque
            # de itens que não foram encontrados automaticamente.
            est_ant = parse_moeda(str(df_bling.at[csv_idx, 'Estoque']).strip()
                                  if 'Estoque' in df_bling.columns else '0')
            est_nov = est_ant + (qtd_estoque or 0)
            if 'Estoque' in df_bling.columns:
                df_bling.at[csv_idx, 'Estoque'] = fmt_qtd(est_nov)
            log_callback(
                f"REASSOC. ESTOQUE [SKU {sku_csv}]: '{nome_perfume}' → {desc_csv} | "
                f"{fmt_qtd(est_ant)} + {fmt_qtd(qtd_estoque or 0)} = {fmt_qtd(est_nov)} un",
                "alerta"
            )
            resultado['preview_rows'].append({
                'sku': sku_csv, 'csv_idx': csv_idx, 'descricao': desc_csv, 'campo': 'Estoque',
                'antes': fmt_qtd(est_ant), 'depois': fmt_qtd(est_nov),
                'gtin': gtin_csv, 'diff_preco': 0,
                'qtd_forn': qtd_estoque,
            })
        else:
            # Modo preço: mesma fórmula da Passada 2
            preco_compra      = valor_unitario * cotacao_dolar
            preco_custo       = preco_compra * (1 + taxa_frete / 100.0)
            preco_venda_final = arredondar_preco(preco_custo * 2)
            desc_forn         = gerar_descricao_fornecedor(nome_perfume, quantidade_ml)
            preco_antigo      = str(df_bling.at[csv_idx, 'Preço']).strip() \
                                if 'Preço' in df_bling.columns else '0,00'
            preco_novo        = formatar_moeda(preco_venda_final)
            diff_preco        = abs(preco_venda_final - parse_moeda(preco_antigo))

            if 'Preço de Compra' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço de Compra'] = formatar_moeda(preco_compra)
            if 'Preço de custo' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço de custo']  = formatar_moeda(preco_custo)
            if 'Preço' in df_bling.columns:
                df_bling.at[csv_idx, 'Preço']            = preco_novo
            if 'Descrição do Produto no Fornecedor' in df_bling.columns:
                df_bling.at[csv_idx, 'Descrição do Produto no Fornecedor'] = desc_forn

            log_callback(
                f"REASSOC. [SKU {sku_csv}]: '{nome_perfume}' → {desc_csv} | "
                f"R$ {preco_antigo} -> R$ {preco_novo}",
                "alerta"
            )
            resultado['preview_rows'].append({
                'sku': sku_csv, 'csv_idx': csv_idx, 'descricao': desc_csv, 'campo': 'Preço',
                'antes': preco_antigo, 'depois': preco_novo,
                'gtin': gtin_csv, 'diff_preco': diff_preco,
                'variacao_pequena': (0 < diff_preco <= limite_variacao),
                'custo': formatar_moeda(preco_custo),
                'compra': formatar_moeda(preco_compra),
                'desc_forn': desc_forn,
                'nome_xlsx': nome_perfume,
                'qtd_forn': qtd_estoque,
            })

        resultado['indices_alterados'].add(csv_idx)
        resultado['ok_count'] = resultado.get('ok_count', 0) + 1
