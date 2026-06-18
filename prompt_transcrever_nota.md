Analise a imagem da nota/pedido do fornecedor e transcreva todos os produtos para uma tabela Excel.

A tabela deve ter exatamente estas colunas, com estes nomes escritos da mesma forma:

| Nome do Perfume | Quantidade (ml) | Valor Unitário | Estoque (un) |

---

REGRAS PARA CADA COLUNA:

Nome do Perfume:
  Nome comercial do produto exatamente como aparece na nota, sem incluir o volume.
  Exemplos: Lattafa Asad | Club de Nuit Intense Man | Orientica Royal Bleu

Quantidade (ml):
  Apenas o número do volume em mililitros, sem unidade.
  Exemplos: 100 | 80 | 125
  Se estiver em FL.OZ., converta para ml multiplicando por 29,57.

Valor Unitário:
  Preço unitário como número decimal com ponto.
  Exemplos: 15.00 | 22.50
  Se a nota mostrar valor total da linha, divida pelo número de unidades compradas.

Estoque (un):
  Quantidade de unidades compradas neste pedido.
  Exemplos: 6 | 12
  Se não aparecer na nota, deixe a célula em branco.

---

ATENÇÃO:
- Não inclua colunas extras, totais, cabeçalhos de seção ou linhas vazias
- Não some linhas de produtos iguais — cada linha da nota vira uma linha da tabela
- Se um mesmo perfume vier em volumes diferentes (ex: 50ml e 100ml), crie uma linha para cada
- Salve o resultado como .xlsx
