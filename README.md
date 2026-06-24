# Automatizador de PPT

Ferramenta para atualizar um PowerPoint mapeado a partir de planilhas de dados.

O modelo atual usa nomes numericos nos graficos do PPT, como `7792738590`, mas os arquivos `.xlsx` dentro do ZIP nao precisam ter esse mesmo nome. Quando o nome bate, o sistema usa isso como atalho. Quando nao bate, ele compara automaticamente colunas, linhas, pergunta da tabela, variavel/abertura do mapeamento e metadados opcionais dentro do XLSX para sugerir o datasource mais provavel.

## Como rodar

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Na interface, envie:

- o arquivo `.pptx` modelo;
- o `.zip` com os datasources.
- no modo com mapeamento, envie tambem a planilha `.xlsx` de mapeamento.

Depois siga o fluxo guiado no app:

0. `Etapa 0 - Projeto`: escolha `Squad1` a `Squad5` e crie ou selecione um projeto.
1. `Etapa 1 - Arquivos`: escolha o modo e envie PPT, ZIP e, se necessario, a planilha de mapeamento.
2. `Etapa 2 - Mapeamento`: confira os pares grafico/datasource e ajuste o datasource escolhido.
3. `Etapa 3 - Dados`: confira a matriz que sera gravada em cada grafico.
4. `Etapa 4 - Validacao`: confira qual arquivo alimenta qual slide/grafico, com alertas de linhas, colunas, valores e confianca.
5. `Etapa 5 - Gerar PPT`: baixe o arquivo atualizado.

A coluna `atualizar_grafico` da planilha pode ser usada como filtro pela interface, mas por padrao todos os graficos encontrados com datasource correspondente ficam selecionaveis.

## Squads, projetos e execucoes

O produto organiza o trabalho assim:

- `Squad1` a `Squad5`: divisao inicial dos times.
- `Projetos`: cada squad cria quantos projetos quiser.
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
- enviar um ZIP com todos os XLSX, mesmo com nomes aleatorios;
- conferir a tela de correspondencias, olhando o score, o contexto do slide e os candidatos quando houver duvida;
- trocar o `Datasource escolhido` diretamente na tela se a sugestao nao estiver correta;
- gerar o PPT.

No modo `Automatico`, a planilha de mapeamento deixa de ser obrigatoria. O sistema olha todos os graficos do PPT, compara com todos os XLSX e monta os melhores pares um-para-um.

Para reforcar o auto-match, um XLSX pode conter nas primeiras linhas pares como `PPT_TAG`, `graph_id`, `var_analise`, `abertura`, `nome_grafico` ou `nome_original`. Isso e opcional; serve apenas como uma pista extra para casos em que duas tabelas sejam muito parecidas.

## IA no mapeamento

O app funciona sem IA, mas quando a chave esta configurada ele usa IA por padrao para revisar o mapeamento. Para habilitar, crie um arquivo `.env` baseado em `.env.example`:

```env
OPENAI_API_KEY=sua_chave
OPENAI_MODEL=gpt-5.5
```

A IA revisa os pares grafico/datasource em lote e preenche `Datasource escolhido` na tela de conferencia. Se a IA discordar da heuristica com baixa confianca, o item fica sinalizado para revisao antes de gerar o PPT.

## Validacao de correspondencias

A etapa `Validacao` substitui a validacao visual. Ela mostra, antes da geracao:

- qual XLSX sera usado em cada slide/grafico;
- pergunta/titulo detectado no XLSX;
- contexto de texto encontrado no slide;
- percentual de linhas e colunas compativeis;
- percentual de valores preenchidos;
- alertas para baixa confianca, colunas/linhas pouco compativeis ou valores vazios.

Se a sugestao estiver errada, a etapa `Mapeamento` permite enviar um XLSX correto e aplicar esse arquivo apenas ao grafico escolhido, sem renomear arquivos.

## Formulas no Excel

A interface sempre tenta calcular formulas antes de ler as planilhas.

O motor tenta primeiro usar o Excel instalado no Windows via `pywin32`, calculando o workbook em uma copia temporaria e substituindo formulas por valores estaticos antes da leitura. Na AWS/Linux ele tenta LibreOffice headless para recalcular os caches do Excel. Se nenhum dos dois estiver disponivel, ele usa um avaliador interno para formulas comuns, incluindo referencias de celulas/ranges, operacoes aritmeticas, `SUM`/`SOMA`, `AVERAGE`/`MEDIA`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`/`SE`, `SUMIF`/`SOMASE` e `COUNTIF`/`CONT.SE`.

Os arquivos originais nao sao alterados.

## Teste rapido

```powershell
python scripts/smoke_test.py
```

Esse teste usa os arquivos de exemplo da pasta, cria um PPT em `outputs/` e tambem valida o auto-match com os datasources renomeados.

## Deploy AWS

O deploy atual foi preparado para ECS Fargate com build na propria AWS via CodeBuild, entao nao precisa de Docker local:

```powershell
.\infra\aws\deploy_fargate.ps1 -AppName qwst-auto-ppt -Region sa-east-1 -AllowedCidr 0.0.0.0/0
```

O script cria/atualiza S3, ECR, CodeBuild, IAM, CloudWatch Logs, ECS Cluster, Security Group, Task Definition e Service. Ele tambem le `OPENAI_API_KEY` do `.env` local e grava no AWS Secrets Manager.

Para pausar o servico e reduzir custo quando nao estiver em uso:

```powershell
.\infra\aws\stop_fargate.ps1 -AppName qwst-auto-ppt -Region sa-east-1
```

Antes de colocar para todos os times, o ideal e trocar o acesso publico por ALB com HTTPS e autenticacao corporativa.

## Estrategia tecnica

- O PPT e lido como pacote Office Open XML.
- Cada grafico nomeado no slide e associado ao `chart.xml` e ao workbook Excel embutido correspondente.
- Os datasources SPSS sao convertidos para a matriz esperada pelo workbook do grafico.
- O gerador atualiza tanto o workbook embutido quanto o cache XML do grafico, preservando o layout visual do template.
