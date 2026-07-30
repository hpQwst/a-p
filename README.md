# Automatizador de PPT

Ferramenta para atualizar um PowerPoint a partir de planilhas de dados.

O core novo trabalha com `PptTarget` generico, nao apenas grafico. Um slide pode ter varios targets atualizaveis, incluindo grafico real do PowerPoint, tabela PowerPoint e, na evolucao, caixas de texto/shapes numericos.

O modelo atual usa nomes numericos nos graficos do PPT, como `7792738590`, mas os arquivos `.xlsx` dentro do ZIP nao precisam ter esse mesmo nome. Quando o nome bate, o sistema usa isso como atalho. Quando nao bate, ele compara automaticamente colunas, linhas, pergunta da tabela, variavel/abertura do mapeamento e metadados opcionais dentro do XLSX para sugerir o datasource mais provavel.

## Como rodar

```powershell
pip install -r requirements.txt
uvicorn web.main:app --host 0.0.0.0 --port 8501
```

Na interface FastAPI, envie:

- o arquivo `.pptx` modelo;
- o `.zip` com os datasources.

O fluxo atual da UI web e:

1. `Acesso`: entre com Microsoft Entra. No primeiro login, usuário comum escolhe
   uma única squad; administrador pode visualizar qualquer uma.
2. `Projeto`: escolha ou crie um projeto dentro da squad permitida.
3. `Arquivos`: envie o PPTX modelo e o ZIP com os XLSX.
4. `Preview`: acompanhe o progresso por objeto, confira os matches e ajuste os
   pendentes. O trabalho é salvo automaticamente e também pelo botão
   `Salvar trabalho`.
5. `Download`: acompanhe a geração por objeto e baixe o PPT atualizado.

## Preparador de Modelo

Para apresentações recorrentes, use `Preparar modelo` na tela inicial:

1. envie o PPTX original e dê um nome ao modelo;
2. o sistema cria uma cópia identificada e mapeia todos os gráficos e tabelas;
3. confira os IDs, títulos detectados, linhas e colunas no estúdio;
4. marque `ativo=1` somente nos objetos que devem mudar;
5. preencha na tela ou baixe o XLSX de mapeamento;
6. envie o XLSX preenchido e todas as planilhas citadas;
7. se o pacote estiver completo, o mapeamento é salvo e o preview abre direto.

O arquivo original nunca é alterado. Objetos com `ativo=0` ficam intactos. Antes
de publicar o modelo, a importação abre cada fonte e bloqueia arquivo ausente,
aba inexistente ou estrutura incompatível. XLSX e ZIP já enviados ficam retidos
no rascunho, então uma correção pode enviar somente o que faltou.

Para tabelas que crescem todo mês:

- `auto`: prefere uma Tabela do Excel; sem ela, detecta o bloco principal;
- `tabela_excel`: usa uma Tabela nomeada, a opção mais segura para crescer sem
  capturar cálculos auxiliares ao lado;
- `dinamico`: expande um intervalo-semente por linhas e colunas contíguas;
- `exato`: nunca sai do intervalo informado.

O estúdio tem salvamento automático, botão `Salvar`, tema claro/escuro e filtros.
O preparador é determinístico e não chama IA. Os modelos ficam isolados por
squad e seus arquivos são versionados; as planilhas de cada rodada continuam no
checkpoint do projeto, sem uma segunda cópia permanente no modelo.

Detalhes do contrato do XLSX: [docs/preparador-modelo.md](docs/preparador-modelo.md).

## Core novo de targets

A nova arquitetura separa as responsabilidades principais:

- `ppt_discovery.py`: descobre targets no PPT, incluindo `chart`, `table`, `text` e `shape`.
- `model_preparer.py`: cria a cópia identificada, o manifesto estrutural, o XLSX de mapeamento e valida o pacote importado.
- `xlsx_parser.py`: interpreta XLSX sem assumir layout fixo antigo.
- `table_normalizer.py`: cria o plano de transformacao e transpõe quando necessario.
- `ai_mapper.py`: monta o payload estrutural para a IA revisar target, datasource e plano.
- `ppt_chart_writer.py`: atualiza chart XML e workbook embutido preservando o grafico.
- `ppt_table_writer.py`: atualiza celulas de tabela PowerPoint preservando estilo.
- `preview_model.py`: gera o modelo amigavel para a UI.
- `engine.py`: orquestra analise, preview e geracao do PPT.

Caso coberto pela regressao MB:

- `3334180514`: chart no slide 1, datasource em series nas linhas e meses nas colunas, transposto para meses nas linhas e series nas colunas.
- `1424058794`: tabela PowerPoint no slide 1, preenchida com uma serie unica formatada em pt-BR.

## Squads, projetos e execucoes

O produto organiza o trabalho assim:

- `Squad1` a `Squad5`: divisao inicial dos times.
- `Projetos`: cada squad cria quantos projetos quiser.
- `Modelos de mapeamento`: memoria reutilizavel por squad, separada dos projetos. Um mapeamento criado no Squad2 nunca aparece como candidato para projetos do Squad1.
- `Execucoes`: cada geracao salva uma nova pasta/objeto com inputs, PPT final e relatorio JSON.
- `Memoria`: correcoes manuais de mapeamento ficam salvas no projeto para auditoria e evolucao futura.
- `Usuários`: perfis comuns ficam isolados em uma squad; administradores alteram
  squad, papel e status em `/admin/users`.

Por padrao, em desenvolvimento, isso fica em `workspace_data/` e nao vai para o Git. Na AWS, o mesmo contrato usa S3 com:

```env
AUTO_PPT_STORAGE_BACKEND=s3
AUTO_PPT_S3_BUCKET=nome-do-bucket
AUTO_PPT_S3_PREFIX=auto-ppt
```

## Mapeamento automatico de datasources

O fluxo mais amigavel e:

- enviar o PPT modelo;
- selecionar um modelo de mapeamento salvo do squad, quando existir, ou deixar o preview sugerir candidatos do proprio squad;
- enviar um ZIP com todos os XLSX, mesmo com nomes aleatorios;
- conferir a tela de correspondencias, olhando o score, o contexto do slide e os candidatos quando houver duvida;
- trocar o `Datasource escolhido` diretamente na tela se a sugestao nao estiver correta;
- gerar o PPT.

Depois de um download bem-sucedido, o sistema cria ou atualiza automaticamente o modelo de mapeamento usado pelo projeto. Na proxima atualizacao, se o PPT tiver os mesmos targets e os datasources mantiverem os mesmos nomes, o template aplica o de-para antes da IA. Targets novos ficam visiveis no preview para serem adicionados ao mesmo modelo.

No modo `Automatico`, a planilha de mapeamento deixa de ser obrigatoria. O sistema olha todos os graficos do PPT, compara com todos os XLSX e monta os melhores pares um-para-um.

Para reforcar o auto-match, um XLSX pode conter nas primeiras linhas pares como `PPT_TAG`, `graph_id`, `var_analise`, `abertura`, `nome_grafico` ou `nome_original`. Isso e opcional; serve apenas como uma pista extra para casos em que duas tabelas sejam muito parecidas.

## IA no mapeamento e na normalizacao

O app funciona sem IA. Mesmo com a chave configurada, nenhuma revisão roda automaticamente por padrão: a pessoa aciona a IA no target ou slide que precisa. Para habilitar, crie um arquivo `.env` baseado em `.env.example`:

```env
OPENAI_API_KEY=sua_chave
OPENAI_MODEL=gpt-5.6-terra
OPENAI_MODEL_SOURCE_MATCH=gpt-5.6-luna
OPENAI_MODEL_SLIDE_MATRIX_BUILDER=gpt-5.6-terra
```

A IA recebe, por target:

- o contrato do PPT extraido do `Editar dados` do grafico ou da tabela PowerPoint;
- um manifesto semantico compacto do XLSX, sem repetir o preview de valores;
- um dump textual compacto das celulas uteis do XLSX, com coordenadas e valores raw;
- o contexto textual estruturado do slide e o nome/contrato OpenXML de cada target.

Com isso ela diagnostica se a acao correta e alinhar, transpor ou pedir revisao. A matriz tecnica continua sendo exibida para o usuario antes do download, e a pessoa pode substituir o XLSX de um target diretamente no card do preview.

Por padrao, o app usa dumps compactos para reduzir custo e latencia sem perder rastreabilidade. Para investigacao pesada, use `AUTO_PPT_AI_XLSX_DUMP_MODE=verbose`. A IA recebe somente estrutura e texto extraídos dos pacotes Office Open XML; imagens de slides não fazem parte do fluxo.

A IA por slide nao roda automaticamente no preview inicial. Isso mantem a tela rapida e evita que uma resposta da IA reestruture graficos que o normalizador deterministico ja mapeou bem. Para investigar um deck dificil, habilite explicitamente `AUTO_PPT_AUTO_SLIDE_AI=1`; para aplicar matrizes geradas por IA no PPT final, habilite tambem `AUTO_PPT_APPLY_SLIDE_AI_OUTPUTS=1`.

No caminho recomendado, a IA de mapeamento recebe registros JSONL por slide: objetos do PPT com colunas/linhas/titulos do `Editar dados` e datasources do mesmo slide com colunas/linhas/titulos detectados. A resposta deve escolher o datasource de cada objeto e pode sugerir uma `recipe_suggestion` estrutural pequena, por exemplo `keep`, `transpose`, `drop_and_keep` ou `drop_and_transpose`. O sistema aplica e valida a transformacao com codigo deterministico; a IA nao devolve nem grava a matriz final.

Essa revisao enxuta de datasource é explícita por padrão (`AUTO_PPT_AI_AUTO_SOURCE_REVIEW=0`). O upload manual e a inclusão de slides nunca acionam IA escondida. Se o time optar por revisão automática de targets sem match ou abaixo do piso de confiança, pode habilitar `AUTO_PPT_AI_AUTO_SOURCE_REVIEW=1`.

Em decks grandes, a revisao enxuta cobre todos os targets pendentes numa unica passada de preview, em lotes de `AUTO_PPT_AI_SOURCE_MATCH_BATCH_TARGETS` (padrao 10) ate o teto de `AUTO_PPT_AI_MATCH_MAX_CALLS` chamadas (padrao 12, ou seja, ate 120 targets por passada). O cache e salvo apos cada lote. A revisao por slide (texto estruturado + matriz) e limitada a `AUTO_PPT_SLIDE_AI_MAX_SLIDES_PER_RUN` slides por execucao (padrao 1), priorizando os slides com targets sem match ou de baixa confianca.

Cada operação de IA usa modelo e esforço próprios. Configuração inicial: `gpt-5.6-luna` com esforço `none` para source match/diagnóstico; `gpt-5.6-terra` com esforço `low` para matriz final. A revisão de matriz faz uma chamada única: seleção da fonte e montagem deixaram de duplicar manifesto, dump e contrato em duas chamadas. O log persistente registra operação, modelo, esforço, bytes, tokens, cache, latência e custo estimado, sem gravar prompt nem resposta.

Antes de subir o servidor, valide a conexao com a OpenAI pelo PowerShell:

```powershell
.\.venv\Scripts\python.exe scripts\check_openai.py
```

Se esse comando retornar `OpenAI: ok`, o app web conseguira usar a IA quando for iniciado pelo mesmo ambiente.

## Validacao de correspondencias

A etapa `Validacao` substitui a validacao visual. Ela mostra, antes da geracao:

- todos os targets descobertos em cada slide, incluindo graficos e tabelas;
- qual XLSX sera usado em cada target;
- o contrato do PPT, equivalente ao que aparece em `Editar dados`;
- a estrutura detectada no XLSX;
- a acao escolhida: alinhar, transpor ou preencher tabela;
- a matriz final que sera gravada no PowerPoint;
- o diagnostico da IA quando a chave estiver configurada.

Se a sugestao estiver errada, o card do target permite enviar um XLSX correto e aplicar esse arquivo apenas ao objeto escolhido, sem renomear arquivos.

## Formulas no Excel

O servidor interpreta formulas sem iniciar Office, COM ou LibreOffice. O avaliador interno cobre referencias de celulas/ranges, operacoes aritmeticas, `SUM`/`SOMA`, `SUMPRODUCT`/`SOMARPRODUTO`, `AVERAGE`/`MEDIA`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`/`SE`, `SUMIF`/`SOMASE`, `COUNTIF`/`CONT.SE` e `SQRT`/`RAIZ`.

Uma formula fora desse contrato interrompe o preview com uma mensagem clara; assim o sistema nao inventa valores. Como excecao consciente, `AUTO_PPT_FORMULA_FALLBACK=cached` permite usar o valor de cache ja salvo pelo autor do XLSX.

Os arquivos originais nao sao alterados.

## Graficos editaveis e Excel embutido

Para preservar o comando `Editar dados` do PowerPoint, o sistema nao usa `python-pptx chart.replace_data()` nem salva o workbook embutido inteiro com `openpyxl.save()` no caminho principal. O caminho validado e serverless: abrir o PPTX/XLSX como pacotes Office Open XML, alterar somente as partes necessarias, preservar a estrutura ZIP/OPC original e atualizar tambem o cache visual do grafico.

Isso significa:

- em desenvolvimento Windows, o app continua funcionando normalmente;
- no container Linux do App Runner, a geracao final funciona sem Microsoft Office;
- o workbook embutido mantem dados completos para `Editar dados`;
- o `chart.xml` e atualizado para o grafico ja abrir visualmente correto;
- tabelas PowerPoint/DrawingML sao atualizadas diretamente no XML preservando estilo.

Essa decisao evita entregar PPT aparentemente correto que depois quebra quando o usuario clica em `Editar dados`.

O preview e a IA trabalham com contratos OpenXML, titulos, contexto textual e dumps estruturados dos XLSX. O sistema nao renderiza imagens de slides nem depende de qualquer aplicativo Office.

## Teste rapido

```powershell
.\.venv\Scripts\python.exe -m unittest discover tests
```

Antes de deployar, rode tambem a regressao com os decks reais. Ela trabalha
somente em copias dentro de `workspace_data/`, compara o pacote Open XML,
workbooks embutidos, tabelas e geometria, abre original e resultado no PowerPoint
e rejeita mudancas visuais fora dos objetos atualizados:

```powershell
.\.venv\Scripts\python.exe .\scripts\run_real_deck_regression.py --include-large --render
```

Por padrao, as fixtures ficam em
`C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos`. Use
`AUTO_PPT_REAL_FIXTURE_ROOT` ou `--fixtures-root` para apontar para outra copia.
Os casos atuais cobrem `andre`, `andre-enxuto`, `hugo`, `mb` e o `mb2` grande.

## Deploy AWS

Produção usa um único serviço AWS App Runner em `us-east-1`, com 1 vCPU, 2 GB,
máximo de uma instância e estado durável no S3. O acesso é exclusivamente pelo
Microsoft Entra; a senha compartilhada está desativada.

Veja [DEPLOYMENT.md](DEPLOYMENT.md) e
[infra/aws/README.md](infra/aws/README.md) para arquitetura, controles de custo,
benchmark de memória e comando de publicação.

## Estrategia tecnica

- O PPT e lido como pacote Office Open XML.
- Cada grafico nomeado no slide e associado ao `chart.xml` e ao workbook Excel embutido correspondente.
- Os datasources SPSS sao convertidos para a matriz esperada pelo workbook do grafico.
- O gerador atualiza tanto o workbook embutido quanto o cache XML do grafico, preservando o layout visual do template.
