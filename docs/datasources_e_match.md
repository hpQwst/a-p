# Datasources e match automático (charts/tabelas do PPT)

Este documento descreve **como os arquivos XLSX devem vir** e **como o sistema
decide qual datasource alimenta cada objeto** (gráfico ou tabela) do PPT. A meta
é que o cliente só precise enviar **o `.pptx` + um `.zip` de `.xlsx`** e o sistema
faça o match perfeito, transpondo linhas/colunas quando necessário.

## TL;DR do formato ideal

- Um `.xlsx` por objeto atualizável do slide.
- **Nome do arquivo** contendo o slide: `..._slide<N>.xlsx` (ex.: `tab39347_slide3.xlsx`).
  O token `slideN` é usado como *blocking* (restringe candidatos àquele slide).
- Primeira linha = **cabeçalho** (rótulos das colunas); primeira coluna = **rótulos das linhas**.
- Valores **verbatim** (exatamente como devem aparecer). Não pré-formate como texto
  com `%` — o percentual/1 casa decimal é aplicado só na exibição (ver "Formatação").
- Opcional, mas o que dá match 100% determinístico: uma célula de **contexto**
  com `obj<numero_do_shape>` (ex.: `obj3958478347`). Os arquivos exportados após
  o primeiro mapeamento já saem assim.

## Como o sistema faz o match (4 camadas)

O match roda em camadas, das mais baratas/certas para as mais caras:

1. **Blocking por slide** — o token `slideN` do nome do arquivo restringe os
   candidatos ao slide correto (`_sources_for_slide`). Arquivos sem `slideN`
   continuam elegíveis para qualquer slide.

2. **Identidade** — se o XLSX carrega `obj<shape>` (em `table_title`/contexto) e
   esse número bate com o shape do objeto no PPT, o match é exato e a confiança é
   máxima, sem IA (`_source_obj_ids` em `table_normalizer.py`).

3. **Conteúdo + atribuição global** — quando não há id, o sistema pontua cada par
   alvo×datasource por conteúdo (categorias, séries, contexto/título) e resolve o
   **casamento ótimo 1:1 do slide** com o algoritmo húngaro (`_hungarian`), em vez
   de escolhas gulosas independentes. Isso impede dois objetos de pegarem o mesmo
   arquivo. A **confiança é calibrada pela margem** para o 2º melhor candidato: um
   vencedor folgado vira alta confiança e dispensa IA.

4. **IA (só no resíduo)** — apenas os objetos que sobraram ambíguos vão para a IA,
   com um payload enxuto (colunas/linhas do datasource + Editar dados do alvo +
   título próximo + contexto). A matriz final ainda é montada e validada por
   código determinístico — a IA só escolhe/sugere, nunca grava valores direto.

## Transposição (linhas × colunas invertidas)

O "Editar dados" de cada gráfico tem uma **orientação-contrato** (quais rótulos
vão nas linhas e quais nas colunas). O XLSX pode vir na orientação oposta. O
sistema detecta isso comparando a cobertura de rótulos nas duas hipóteses
(`_best_axis_alignment`: `same` vs `cross`) e, se estiverem cruzados, **transpõe**
a matriz para bater exatamente com o contrato do Editar dados antes de gravar.
Ou seja: **você não precisa arrumar a orientação do XLSX manualmente**.

Exemplo (`tests/test_transpose_alignment.py`): contrato com séries nas linhas e
XLSX com séries nas colunas → resultado transposto para o formato do Editar dados.

## Formatação dos números

- O valor gravado é **verbatim** (precisão total) — o "Editar dados" preserva
  todas as casas decimais.
- A **exibição** usa 1 casa decimal. O símbolo `%` só aparece se o **template do
  gráfico** já usar formato percentual; o sistema **nunca inventa** `%` que não
  existe no XLSX nem no contrato do PPT (`effective_value_format` em
  `ppt_chart_writer.py`).

## Aprendizado ("perfeito na 2ª vez")

Depois de um download bem-sucedido, o sistema salva no **template de mapeamento do
squad**, para cada objeto:

- `target_fingerprint` — impressão digital de conteúdo do objeto (tipo +
  categorias + séries), **estável mesmo que o número do shape mude** numa nova
  versão do deck;
- `source_signature` + `source_categories`/`source_series` — assinatura de
  conteúdo do datasource usado, **resistente a renomeação do arquivo**.

Na execução seguinte (`resolve_learned_matches` em `learned_mapping.py`), o sistema
reaplica esse mapeamento **antes** de conteúdo/IA:

- reencontra o objeto por id/alias e, se o deck foi recriado, por `target_fingerprint`;
- reencontra o datasource por nome; se foi renomeado, por `source_signature`; e se
  houve pequena variação, pelo maior *overlap* de rótulos (Jaccard ≥ 0,6).

Resultado comprovado com os arquivos reais do André: mesmo renomeando **todos** os
datasources para nomes genéricos, o sistema reconecta 4/4 por conteúdo.

## Checklist para o cliente

- [ ] Um `.xlsx` por gráfico/tabela a atualizar.
- [ ] Nome do arquivo com `slide<N>`.
- [ ] Cabeçalho na 1ª linha e rótulos na 1ª coluna.
- [ ] Valores numéricos como número (sem formatar como texto/`%`).
- [ ] (Opcional) célula de contexto `obj<shape>` para match exato imediato.
- [ ] Orientação livre — o sistema transpõe se preciso.
