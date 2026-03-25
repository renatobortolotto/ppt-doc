# Guia de Estilo dos Graficos

Este arquivo registra decisoes visuais que ja foram aprovadas no projeto e devem
ser reutilizadas em ajustes futuros.

## Slide 11 - Despesas empilhadas

Referencia atual:

- `utils/slide11_charts.py`
- funcoes: `_plot_stacked_expenses` e `generate_slide11_charts`

Padrao aprovado:

- fundo transparente
- barras empilhadas mais finas do que o default original
- barras mais proximas entre si
- valores internos centralizados em cada faixa
- total acima de cada barra
- total da ultima barra em negrito
- bracket de variacao sempre acima dos totais, nunca no meio do grafico
- nomes das series no lado esquerdo, fora do grafico, alinhados com a altura da
  faixa correspondente da primeira barra
- esses nomes devem ser texto simples, sem quadrado solto, sem badge, sem bloco
  colorido de fundo

Detalhe importante da legenda lateral:

- `Depreciacao e Amortizacao` deve ficar alinhado ao centro da faixa clara da
  primeira barra
- `Administrativas` deve ficar alinhado ao centro da faixa intermediaria da
  primeira barra
- `Pessoal` deve ficar alinhado ao centro da faixa escura da primeira barra

Regra de dados para esse slide:

- despesas podem vir negativas no Excel contabil
- o grafico deve usar a magnitude positiva desses valores
- ou seja, antes de plotar, aplicar valor absoluto nas series de despesa

## Como reaplicar esse estilo

Se um grafico novo precisar seguir esse mesmo modelo:

- usar barras estreitas com espacamento comprimido
- posicionar os rotulos de serie por coordenada vertical da primeira barra
- evitar legenda tradicional separada do grafico
- manter labels e brackets legiveis dentro do espacamento compacto

## Observacao

Se houver conflito entre esse guia e um pedido visual novo do usuario, vale o
pedido mais recente do usuario.
