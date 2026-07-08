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

1. `Projeto`: escolha o squad e informe o nome do projeto.
2. `Arquivos`: envie o PPTX modelo e o ZIP com os XLSX.
3. `Preview`: confira todos os targets descobertos por slide, com tipo, datasource, acao e matriz final. Se algum match estiver errado, envie um XLSX diretamente no card daquele target.
4. `Download`: baixe o PPT atualizado.

## Core novo de targets

A nova arquitetura separa as responsabilidades principais:

- `ppt_discovery.py`: descobre targets no PPT, incluindo `chart`, `table`, `text` e `shape`.
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

O app funciona sem IA, mas quando a chave esta configurada ele usa IA por padrao na etapa de preview. Para habilitar, crie um arquivo `.env` baseado em `.env.example`:

```env
OPENAI_API_KEY=sua_chave
OPENAI_MODEL=gpt-5.5
```

A IA recebe, por target:

- o contrato do PPT extraido do `Editar dados` do grafico ou da tabela PowerPoint;
- um manifesto semantico compacto do XLSX, incluindo titulo da tabela, filtro/label lateral, categorias, series e preview;
- um dump textual compacto das celulas uteis do XLSX, com coordenadas e valores raw;
- a matriz final proposta pelo normalizador;
- o contexto textual do slide, o nome do shape e, na revisao visual por slide, uma imagem otimizada do slide com IDs.

Com isso ela diagnostica se a acao correta e alinhar, transpor ou pedir revisao. A matriz tecnica continua sendo exibida para o usuario antes do download, e a pessoa pode substituir o XLSX de um target diretamente no card do preview.

Por padrao, o app evita mandar dumps verbosos para reduzir custo e latencia sem perder rastreabilidade. Para investigacao pesada, use `AUTO_PPT_AI_XLSX_DUMP_MODE=verbose`. A imagem vai na primeira chamada visual por slide; a segunda chamada de montagem de matriz nao reenvia imagem por padrao. Para habilitar, use `AUTO_PPT_AI_IMAGE_IN_MATRIX=1`. O limite padrao da imagem e `AUTO_PPT_AI_IMAGE_MAX_SIDE=1400`.

A IA por slide nao roda automaticamente no preview inicial. Isso mantem a tela rapida e evita que uma resposta da IA reestruture graficos que o normalizador deterministico ja mapeou bem. Para investigar um deck dificil, habilite explicitamente `AUTO_PPT_AUTO_SLIDE_AI=1`; para aplicar matrizes geradas por IA no PPT final, habilite tambem `AUTO_PPT_APPLY_SLIDE_AI_OUTPUTS=1`.

No caminho recomendado, a IA de mapeamento recebe registros JSONL por slide: objetos do PPT com colunas/linhas/titulos do `Editar dados` e datasources do mesmo slide com colunas/linhas/titulos detectados. A resposta deve escolher o datasource de cada objeto e pode sugerir uma `recipe_suggestion` estrutural pequena, por exemplo `keep`, `transpose`, `drop_and_keep` ou `drop_and_transpose`. O sistema aplica e valida a transformacao com codigo deterministico; a IA nao devolve nem grava a matriz final.

Essa revisao enxuta de datasource roda automaticamente para targets sem match e para matches deterministicos abaixo de `AUTO_PPT_AI_REVIEW_CONFIDENCE_FLOOR` (padrao `0.80`), quando `OPENAI_API_KEY` estiver configurada. Para desligar essa revisao automatica e manter apenas o fluxo deterministico/manual, use `AUTO_PPT_AI_AUTO_SOURCE_REVIEW=0`. Esse mecanismo nao habilita a IA por slide nem a gravacao de matrizes finais por IA.

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

A interface sempre tenta calcular formulas antes de ler as planilhas.

O motor tenta usar o melhor calculador disponivel no ambiente. Em desenvolvimento Windows, pode usar Excel via `pywin32`. Na AWS/Linux, usa LibreOffice headless para recalcular os caches do Excel. Se nenhum dos dois estiver disponivel, ele usa um avaliador interno para formulas comuns, incluindo referencias de celulas/ranges, operacoes aritmeticas, `SUM`/`SOMA`, `AVERAGE`/`MEDIA`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`/`SE`, `SUMIF`/`SOMASE` e `COUNTIF`/`CONT.SE`.

Os arquivos originais nao sao alterados.

## Graficos editaveis e Excel embutido

Para preservar o comando `Editar dados` do PowerPoint, o sistema nao usa `python-pptx chart.replace_data()` nem salva o workbook embutido inteiro com `openpyxl.save()` no caminho principal. O caminho validado e serverless: abrir o PPTX/XLSX como pacotes Office Open XML, alterar somente as partes necessarias, preservar a estrutura ZIP/OPC original e atualizar tambem o cache visual do grafico.

Isso significa:

- em desenvolvimento Windows, o app continua funcionando normalmente;
- em Linux/Docker/Fargate Linux, a geracao final de PPT com graficos editaveis deve funcionar sem Microsoft Office;
- o workbook embutido mantem dados completos para `Editar dados`;
- o `chart.xml` e atualizado para o grafico ja abrir visualmente correto;
- tabelas PowerPoint/DrawingML sao atualizadas diretamente no XML preservando estilo.

Essa decisao evita entregar PPT aparentemente correto que depois quebra quando o usuario clica em `Editar dados`.

Para preview visual e IA por slide, o sistema tenta renderizar o slide com PowerPoint COM quando estiver em Windows local. Em container Linux, usa LibreOffice headless para converter o PPTX em PDF e PyMuPDF para gerar o PNG do slide com os IDs desenhados.

## Teste rapido

```powershell
python scripts/smoke_test.py
python -m unittest tests.test_mb_update_targets
```

O smoke test usa os arquivos de exemplo da pasta, cria um PPT em `outputs/` e valida o auto-match com os datasources renomeados.

O teste MB usa, por padrao, `C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\mb` e valida descoberta de chart+tabela, normalizacao/transposicao, escala de percentuais e escrita no PPT final.

## Deploy AWS

O deploy atual foi preparado para ECS Fargate com build na propria AWS via CodeBuild, entao nao precisa de Docker local. A regiao padrao do projeto e `us-east-1`, e os recursos criados pelo script recebem a tag `Name=qwst-auto-ppt`:

```powershell
.\infra\aws\deploy_fargate.ps1 -AppName qwst-auto-ppt -Region us-east-1 -AllowedCidr 0.0.0.0/0
```

O script cria/atualiza S3, ECR, CodeBuild, IAM, CloudWatch Logs, ECS Cluster, Security Group, Task Definition e Service. Ele tambem le `OPENAI_API_KEY` do `.env` local e grava no AWS Secrets Manager.

Para pausar o servico e reduzir custo quando nao estiver em uso:

```powershell
.\infra\aws\stop_fargate.ps1 -AppName qwst-auto-ppt -Region us-east-1
```

Antes de colocar para todos os times, o ideal e trocar o acesso publico por ALB com HTTPS e autenticacao corporativa.

## Estrategia tecnica

- O PPT e lido como pacote Office Open XML.
- Cada grafico nomeado no slide e associado ao `chart.xml` e ao workbook Excel embutido correspondente.
- Os datasources SPSS sao convertidos para a matriz esperada pelo workbook do grafico.
- O gerador atualiza tanto o workbook embutido quanto o cache XML do grafico, preservando o layout visual do template.
