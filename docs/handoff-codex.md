# Handoff para o Codex — automatizador-ppt

Documento de contexto para continuar o trabalho a partir daqui. Foi escrito por
um agente anterior (Claude) que não deixa histórico acessível ao Codex, então
tudo que importa está aqui.

> **Regra de ouro deste handoff:** os itens pendentes na seção 4 têm **perguntas
> abertas**. NÃO decida sozinho. Levante cada pergunta com o Hugo, colete as
> respostas e só então implemente. O Hugo pediu explicitamente isso.

---

## 0. Como trabalhar neste repositório

1. **graphify — use sempre.** Existe um grafo de conhecimento em `graphify-out/`.
   - Para qualquer pergunta sobre o código, rode primeiro: `graphify query "<pergunta>"`.
   - Relações entre símbolos: `graphify path "<A>" "<B>"`; conceito focado: `graphify explain "<X>"`.
   - Isso devolve um subgrafo pequeno, melhor que grep cego ou ler o `GRAPH_REPORT.md` inteiro.
   - **Depois de alterar código, rode `graphify update .`** para manter o grafo em dia (é AST puro, sem custo de API).

2. **caveman ultra — modo de saída.** Responda terso, estilo "caveman inteligente":
   sem artigos, sem enrolação, fragmentos ok, cada fato uma vez. Mantenha
   verbatim: código, nomes de função/API, comandos e mensagens de erro exatas.
   Largue o caveman em avisos de segurança e confirmações de ação irreversível.

3. **Fronteira de recursos na AWS (conta 134164930693).** Só criar/alterar
   recursos cujo nome comece com `squad4`/`squad5`. Tudo o mais na conta é de
   outras pessoas — não alterar, deletar nem redeployar. Os recursos deste
   projeto são `squad4e5-auto-ppt` (App Runner, ECR, bucket S3, secrets).
   Leitura para diagnóstico (`describe`, `list`, `get`) é permitida.

4. **Repositórios — os dois.** Azure DevOps é a fonte da verdade; GitHub é
   espelho. Todo push vai para os dois:
   ```
   git push azure main
   git push origin main
   ```
   `azure` = `https://dev.azure.com/qwst-equipe-tecnica/qwst-equipe-tecnica/_git/qwst-auto-ppt`
   `origin` = `https://github.com/hpQwst/a-p.git`

5. **Testes.** `unittest`, sem pytest. Rodar com o Python do venv:
   ```
   .venv\Scripts\python.exe -m unittest discover tests
   ```
   Alguns testes dependem de decks reais que ficam **fora** do repositório e se
   auto-pulam (`skipUnless`) quando ausentes. Não fabrique fixtures.

6. **Verifique de verdade.** O padrão desta base é provar, não afirmar: rodar
   contra dados reais, medir, tirar screenshot do navegador. Vários bugs desta
   rodada foram pegos porque o teste rodou de fato (o `os.replace` no Windows, o
   escopo reservado da MSAL, o falso positivo do detector de tabelas).

---

## 1. O que é o projeto

Ferramenta web que atualiza decks PowerPoint a partir de dados de planilha. O
usuário sobe um `.pptx` modelo + planilhas `.xlsx`; o sistema descobre os alvos
atualizáveis (gráficos e tabelas), casa cada um com a planilha certa, normaliza
os dados e grava um `.pptx` novo que ainda abre no PowerPoint (inclusive "Editar
dados" nos gráficos).

O núcleo é **cirurgia OOXML sem Office**: abre `.pptx`/`.xlsx` como ZIP/OPC,
edita só as partes necessárias, atualiza o cache visual do gráfico. Isso é o que
permite rodar em container Linux headless. **Não reintroduza Office/COM/LibreOffice.**

Entrada atual: `web/main.py` (FastAPI + Jinja). `app.py` (Streamlit) é legado —
não estenda.

Organização: **Squads** (squad1–squad5) → Projetos → Modelos de mapeamento
(memória de mapeamento por squad) → Execuções.

Arquivos-chave:
- `web/main.py` — rotas, middleware de auth, geração, preview.
- `web/auth.py` — sessão, config Entra + senha da equipe, `current_user`, `session_subject`.
- `web/entra.py` — fluxo OIDC Microsoft, `exchange_code` (checagem de tenant).
- `web/audit.py` — quem-fez-o-quê (`actor_from`, `record`, `remember_actor`).
- `ppt_automator/project_store.py` — storage (local/s3), escrita atômica + lock, templates, runs, correções.
- `ppt_automator/engine.py` — `analyze_update_package`, `generate_updated_pptx` (onde entraria instrumentação de progresso).
- `worker/processor.py` — `analyze_files`, `_analysis_warnings` (travas de estrutura de planilha).
- `CLAUDE.md` — guia da base (mantê-lo atualizado).

---

## 2. Estado do deploy

- **AWS App Runner**, região `us-east-1`, serviço `squad4e5-auto-ppt`, 1 vCPU / 2 GB, **máx. 1 instância** (deliberado — ver abaixo).
- URL: `https://sjsgq73and.us-east-1.awsapprunner.com`
- Estado compartilhado no S3 (`AUTO_PPT_STORAGE_BACKEND=s3`, bucket `squad4e5-auto-ppt-<conta>`, versionamento ligado).
- Build: código empacotado por `git archive` → zip no S3 → CodeBuild → ECR. **CodeBuild não lê Azure Repos**, por isso o caminho por zip.
- Deploy: `.\infra\aws\deploy.ps1 -TeamPassword "..." [-EntraTenantId ... -EntraClientId ... -EntraClientSecret ... -EntraRedirectUri ...]`
- Infra: `infra/aws/apprunner.yaml` + `infra/aws/build.yaml` + `infra/aws/deploy.ps1`. Runbook em `DEPLOYMENT.md`.

**Por que máx. 1 instância:** o `project_store` grava JSON no S3; cada gravação é
atômica, mas duas instâncias gravando o mesmo objeto poderiam perder uma
atualização. Aumentar `MaxSize` exige antes escrita condicional por ETag.

**Login Microsoft (Entra OIDC):** já foi deployado e **confirmado funcionando**
com a conta `hugo.rocha@qwst.co`. Single-tenant (`tid` conferido em
`entra.exchange_code`). Senha da equipe existe como reserva. Valores:
- `ENTRA_TENANT_ID=6d620aff-4c64-4458-bac3-2e502b255ee1`
- `ENTRA_CLIENT_ID=9bf49a4b-2d65-4973-943d-d57e1c60c3a0`
- `ENTRA_REDIRECT_URI=https://sjsgq73and.us-east-1.awsapprunner.com/auth/callback`
- Client secret e `AUTO_PPT_SESSION_SECRET`: no Secrets Manager. **Não regenerar o
  session secret** — derrubaria todas as sessões ativas.

> **ATENÇÃO — pode haver commits ainda não deployados.** Os últimos commits
> (nome do PPT de saída, geração em segundo plano por padrão, trilha de
> auditoria) foram para o git DEPOIS do último deploy confirmado. Primeira coisa
> a fazer: comparar a imagem que está no App Runner com o `HEAD` e, se diferente,
> rodar `deploy.ps1` para publicar. Comando de checagem:
> ```
> aws apprunner describe-service --service-arn <arn> --region us-east-1 --query "Service.SourceConfiguration.ImageRepository.ImageIdentifier"
> ```

---

## 3. O que foi feito nesta rodada (para contexto, não refazer)

Em ordem, do mais antigo ao mais recente:
1. Sistema visual "Bancada" + 7 mudanças de UX zero-treino (L1–L7): prévia
   honesta (tabela + gráfico fiel com tipo/cores reais do XML), modo avançado que
   esconde diagnóstico, gate de conclusão, escolha de planilha por clique, entrada
   de vários `.xlsx` soltos, onboarding de 3 passos, linguagem de usuário.
2. Escritas atômicas + lock no `project_store` (evita corromper mapeamento). Bug
   do `os.replace` no Windows corrigido (antivírus segura `.pptx`).
3. Fargate removido → App Runner. `boto3` mantido (backend S3).
4. Build por zip no S3. Dockerfile puxa imagem base do espelho AWS (`public.ecr.aws/docker/library/python:3.12-slim`) por causa de `429 Too Many Requests` do Docker Hub.
5. Travas de estrutura de planilha: avisa aba ignorada, mais de uma tabela na
   mesma aba, nomes de arquivo repetidos (recusa). Detector **conservador** (só
   linha em branco separa; validado em dados reais, zero falso positivo).
6. Login Microsoft Entra (OIDC) + senha da equipe como reserva.
7. Trilha de auditoria (`web/audit.py`): grava quem gerou, trocou planilha,
   treinou memória, aprovou/pulou. Com senha compartilhada grava
   `identified: false` em vez de inventar nome.
8. Mantém o nome do PPT enviado + sufixo `__<data>_<hora>`. Geração em segundo
   plano ligada por padrão (`AUTO_PPT_ASYNC_GENERATION=1`).

---

## 4. Trabalho pendente — LEVANTE AS PERGUNTAS COM O HUGO, NÃO DECIDA

### A. Barra de porcentagem REAL na geração

**Estado hoje:** a geração roda em segundo plano; o overlay mostra tempo
decorrido, mas **sem porcentagem** — a geração não reporta progresso por slide, e
uma barra inventada mentiria. O Hugo quer barra real.

**Onde mexer:** `generate_updated_pptx` / `engine.py` percorre os alvos; dá para
emitir progresso via callback que escreve em `generation_processing.json`
(`_save_generation_state`), e o frontend (`app.js`, `updateGenerationProgress`)
lê no polling e desenha a barra. O **preview** já tem progresso por slide em
`preview_processing.json` — reusar esse padrão.

**Perguntas para o Hugo:**
- Granularidade: barra por **slide** ou por **objeto/target**? (Deck de 100
  slides com N objetos cada — qual conta o usuário quer ver avançar?)
- Quer barra também no **preview** (hoje só status textual) ou só na geração?
- Mostrar contagem ("47 de 120 gráficos") junto da barra, ou só a barra?

### B. Item 1 — jobs em `/tmp` são efêmeros

**Estado hoje:** jobs vivem em `RUNTIME_ROOT` (`/tmp/auto-ppt-jobs` no App Runner).
O checkpoint no S3 só é salvo quando o preview **conclui**. Se a instância
reciclar no meio de uma análise, aquele job em andamento se perde.

**Contexto que atenua:** 1 instância, uso de 3–4×/mês. Reciclagem bem no meio de
um uso é improvável — talvez ROI baixo para uma solução cara.

**Opções (para o Hugo escolher, com Codex explicando custo de cada):**
- Persistir o job em andamento no S3 continuamente (sobe input.pptx +
  datasources.zip + estado a cada passo). Custo: escritas S3 frequentes.
- Só recuperação graciosa no restart (retomar de onde o S3 tem, refazer o resto).
- Aceitar o risco e não fazer nada (dado o volume de uso).

**Perguntas para o Hugo:**
- Qual a tolerância? É aceitável o usuário reenviar os arquivos se (raramente) a
  instância reciclar no meio?
- Vale o custo de escritas contínuas no S3 para um evento raro?

### C. Item 2 — limite de memória / tamanho para deck grande

**Estado hoje:** 2 GB de RAM, **não testado** com deck de 100+ slides. Já existem
`AUTO_PPT_MAX_UPLOAD_MB=250` e `AUTO_PPT_MAX_REQUEST_MB=600`, mas são limites de
**upload**, não de **uso de memória**.

**Sugestão do Hugo:** avisar o tamanho máximo combinado de PPT+XLSX; somar
conforme o usuário sobe e checar se estourou.

**O que o Codex deve responder ao Hugo (ele pediu "o que o Codex sugere?"):**
- **Cuidado:** tamanho de arquivo prevê mal o pico de memória. Um `.xlsx` de 10 MB
  pode explodir para muito mais em RAM ao abrir no openpyxl. Somar tamanho de
  arquivo e comparar com um limite fixo pode barrar deck que caberia, e liberar
  deck que estoura.
- **Recomendação de método:** MEDIR primeiro — rodar um deck grande real e
  observar o pico de RAM — antes de escolher qualquer limite. Só então decidir
  entre: (a) aviso de tamanho combinado no upload (o que o Hugo sugeriu, fácil,
  aproximado); (b) aumentar a memória do App Runner (2→4 GB, +custo, sempre
  ligado); (c) limitar nº de slides/objetos por execução; (d) processar em
  streaming.
- **Custo é restrição dura** (o Hugo enfatizou não estourar): subir memória do App
  Runner encarece a instância que fica 24/7. O aviso de tamanho é grátis, mas é
  paliativo, não garantia.

**Perguntas para o Hugo:**
- Tem um deck real grande (100+ slides) para o Codex medir o pico de memória?
- Qual o maior caso real esperado (slides × objetos por slide)?
- Prefere aviso barato-mas-aproximado, ou medir e dimensionar a memória certa
  (com o custo que vier)?

### D. Isolamento por squad + primeiro login + tela de admin

**Decisão do Hugo (já tomada, não reabrir):**
- É **isolamento de verdade**, não só preferência visual.
- No primeiro acesso, a pessoa **seleciona o squad** que ela vê — e fica fixo
  para sempre.
- O login do Hugo (`hugo.rocha@qwst.co`) tem uma **tela a mais** de admin: mudar a
  visualização de alguém, desativar usuário, promover alguém a admin, etc.

**Estado hoje:** o middleware só checa se está autenticado. NÃO existe mapa
usuário→squad, nem isolamento, nem admin. A home lista todos os squads; jobs
carregam `metadata.project.squad` mas ninguém valida dono no acesso.

**Implicações que o Codex precisa cobrir (visual e por trás):**
- **Armazenamento** do mapa usuário→squad e do papel (admin/comum): onde? Novo
  store no `project_store` (ex. `users/<email>.json` no S3)? Como semear
  `hugo.rocha@qwst.co` como admin no primeiro boot?
- **Autorização por requisição** (o furo maior): isolar de verdade exige checar o
  squad do usuário em `/projects/{squad}/...`, `/jobs/{id}/...` (job pertence a um
  squad via metadata), no POST `/preview` (cria projeto num squad) e nas listagens
  da home (`projects_by_squad`, `mapping_templates_by_squad` devem filtrar só o
  squad do usuário). Sem isso, acesso direto por URL fura o isolamento.
- **Fluxo de primeiro login:** depois do callback Entra, se o usuário não tem
  squad → redireciona para uma tela de escolha de squad → grava → trava.
- **Tela de admin:** listar usuários, ver/trocar squad de cada um, desativar
  (bloquear login no middleware), promover/rebaixar admin.
- **Visual:** para o usuário comum, o seletor de squad na home some (ele está
  preso ao seu). Admin vê tudo? Ou admin também escolhe uma visão?

**Perguntas para o Hugo (todas precisam de resposta antes de codar):**
1. Quais squads existem para escolher? Os fixos `squad1`–`squad5`? (O recurso AWS
   se chama `squad4e5` mas os squads lógicos são 5.)
2. `squad4` e `squad5` são **separados** (memórias e projetos isolados entre si) ou
   **compartilham** (já que o recurso é "squad4e5")? Isso muda o modelo de dados.
3. Como o admin é identificado no código? Lista de e-mails admin em
   env/secret, ou flag no store de usuários semeada com `hugo.rocha@qwst.co`?
4. "Desativar usuário" = bloquear login? O que a pessoa vê ao ser bloqueada?
5. O admin pode **trocar** o squad de alguém (o Hugo disse que sim) — mas o
   usuário comum NUNCA pode trocar o seu, certo? Confirmar que "fixo para sempre"
   vale só para o comum.
6. Admin enxerga **todos** os squads ao usar a ferramenta, ou tem uma visão
   neutra + a tela de administração à parte?
7. Com isolamento por squad + auditoria nominal, a **senha da equipe** vira um
   furo (entra sem identidade e sem squad). Desligar a senha em produção? Ou
   deixar só para o admin, em emergência?
8. Precisa de log de quem-mudou-o-quê na própria administração (admin trocou
   squad de fulano)? A trilha de auditoria já existe e pode receber esses eventos.

### E. Coisas que o Hugo pode ter esquecido (o Codex deve levantar também)

1. **Redeploy pendente** (ver seção 2): confirmar imagem no App Runner vs. `HEAD`
   e publicar os últimos commits.
2. **Senha da equipe passou pelo chat** numa mensagem anterior. Se for mantê-la,
   trocá-la: `deploy.ps1 -TeamPassword "nova" -SkipBuild`. (Relacionado à pergunta
   D.7.)
3. **Custo:** ativar a tag de rateio `Name` no Billing para ver o custo só deste
   app, separado do resto da conta (Cost Explorer atrasa ~24h). Considerar
   `budget`/alarme, e `pause-service` se o uso for muito esporádico. App Runner
   custa ~US$10/mês (referência real da conta) mesmo parado, porque a memória fica
   provisionada 24/7.
4. **`CLAUDE.md` e docs:** manter atualizados a cada mudança estrutural (o padrão
   da base é doc bater com código).
5. **Sinalizar reduce-motion / acessibilidade** ao mexer em UI nova (barra de
   progresso, telas de squad/admin).
6. Rodar a suíte inteira (`unittest discover tests`) e `graphify update .` antes de
   cada commit; empurrar para os dois remotos.

---

## 5. Ordem sugerida (o Codex pode propor outra, mas confirme com o Hugo)

Primeiro colete as respostas das perguntas da seção 4. Depois, sugestão de
prioridade:
1. **Redeploy** (E.1) — publicar o que já está pronto.
2. **Isolamento por squad + admin** (D) — é o maior e destrava o resto (auditoria
   só vira nominal de verdade quando cada um tem identidade+squad).
3. **Barra de porcentagem** (A) — melhoria visual de alto valor, escopo contido.
4. **Item 2 memória** (C) — medir com deck real e dimensionar.
5. **Item 1 jobs efêmeros** (B) — provavelmente o de menor ROI; decidir com dado.
