# Handoff 2 — gargalo de matching em deck grande

Continuação do trabalho no `automatizador-ppt`. O handoff anterior é
`docs/handoff-codex.md` (contexto geral do projeto); este aqui trata do problema
específico que ficou aberto.

---

## Regras de trabalho (valem sempre)

1. **graphify sempre.** `graphify query "<pergunta>"` antes de sair lendo código;
   `graphify path "<A>" "<B>"` para relações; `graphify explain "<X>"` para
   conceito focado. Depois de alterar código, `graphify update .`.
2. **caveman ultra** na saída: terso, sem artigos, sem enrolação, cada fato uma
   vez. Verbatim em código, nomes de API, comandos e mensagens de erro. Largue o
   caveman em aviso de segurança e confirmação de ação irreversível.
3. **AWS (conta 134164930693): só toque em recursos `squad4*`/`squad5*`.** O
   resto é de outras pessoas. Este projeto é `squad4e5-auto-ppt`. Leitura
   (`describe`/`list`/`get`) é livre.
4. **Dois remotos.** `git push azure main` **e** `git push origin main`. Azure é a
   fonte da verdade.
5. **Testes:** `.venv\Scripts\python.exe -m unittest discover tests`. Hoje: **137
   passando**. Alguns testes pulam sozinhos sem fixtures externas — normal.
6. **Meça, não afirme.** Todo número abaixo foi medido nos decks reais. Faça o
   mesmo antes e depois de qualquer mudança.

---

## Estado da árvore: HÁ TRABALHO NÃO COMMITADO

`git status` mostra 4 arquivos modificados (+169 linhas), **de propósito não
commitados**:

```
ppt_automator/xlsx_parser.py
ppt_automator/engine.py
worker/processor.py
tests/test_xlsx_structure_guards.py
```

### O que essa mudança faz

**Cada aba de um workbook virou uma fonte independente.**

Motivo: o deck real da TIM tem **639 objetos** e a planilha que o alimenta tem
**178 abas**. O parser lia só a primeira aba, então casava **0 de 639**. Era o
maior buraco funcional do produto.

Desenho adotado (mantenha):
- Workbook com **1 aba** continua com o identificador antigo (o caminho no ZIP).
  Isso preserva os mapeamentos já aprendidos, que casam a fonte pelo nome do
  arquivo.
- Workbook com **N abas** gera N fontes, identificadas como `arquivo.xlsx#Aba`.
- Helpers novos em `xlsx_parser.py`: `zip_path_of()`, `source_sheet_of()`,
  `workbook_sheet_names()`, `_parse_worksheet()`.
- `engine._build_slide_aware_plans` compara escopo por **arquivo** (usando
  `zip_path_of`), porque o escopo por slide é por arquivo, não por aba.
- O aviso antigo "só a primeira aba foi lida" virou informativo ("abertas aba por
  aba, cada uma como uma fonte").

**Testes: 137 passando com a mudança.**

### Por que não foi commitado

Num workbook de 178 abas o **matching** passa de 20 minutos. Antes, o usuário
recebia lixo instantâneo (0 casados); agora esperaria meia hora. Os dois são
inaceitáveis, mas travar 30 min queima CPU e parece pane. **Não commite antes de
resolver o gargalo** — ou commite junto com a solução.

---

## O gargalo — sua tarefa

### Medições (feitas nos decks reais, em `workspace_data/benchmarks/real_decks_2026_07_29/`)

**Deck TIM** (`tim/`): `C_Ap_Completa_Experiencia_2Q26_v01.pptx` 213 MB, **639
alvos**; `TIM_EXP_TB_COMPLETA_a2026m06_v01.xlsx` 858 KB, **178 abas**.

| Etapa | Antes | Depois da correção de parse |
| --- | --- | --- |
| Ler as 178 abas | > 10 min (estourava) | **16,2 s** |
| Matching 639 × 178 | — | **23 a 48 min (projetado)** |

Escala do matching medida com `build_transform_plans`:

```
1 alvo  x 178 fontes: 4,5 s   -> projeção 639 alvos = 48 min
2 alvos x 178 fontes: 4,3 s   -> projeção 639 alvos = 23 min
4 alvos x 178 fontes: 11,5 s  -> projeção 639 alvos = 31 min
```

### Pista mais forte

**1 alvo custa 4,5 s e 2 alvos custam 4,3 s** — praticamente o mesmo. Isso indica
**custo fixo por chamada** de `build_transform_plans`, quase independente do
número de alvos. Suspeita: as features de cada fonte (normalização de rótulos,
`_source_obj_ids`, `_source_axes`, eixos/orientação) são recalculadas a cada
chamada, para as 178 fontes.

Como `_build_slide_aware_plans` hoje chama uma vez **por slide**, um deck de ~118
slides paga esse custo fixo ~118 vezes.

Já foi feito: agrupar alvos por slide (antes era uma chamada **por alvo**, pior
ainda). Ajudou, mas não resolveu.

### O que investigar

1. **Confirme a hipótese com perfil de verdade** (`cProfile`), não por leitura.
   Onde exatamente vão os 4 s fixos? `ppt_automator/table_normalizer.py` é o
   ponto de partida (`build_transform_plans`, `source_match_candidates`,
   `_source_obj_ids`, `_source_axes`).
2. **Traga mais de uma solução**, com custo e risco de cada. Algumas direções —
   não se limite a elas:
   - Pré-computar as features de cada fonte **uma vez por análise** e reusar
     (índice passado adiante, em vez de recalcular por chamada).
   - Pré-filtro barato antes do escore caro (ex.: candidato precisa compartilhar
     ao menos um rótulo), reduzindo o par-a-par.
   - Indexação invertida rótulo → fontes, para não varrer as 178 por alvo.
   - Teto de candidatos por alvo, se houver critério honesto de corte.
3. **Cuidado para não trocar velocidade por qualidade de match.** O algoritmo é
   uma atribuição 1:1 (Hungarian) com confiança calibrada pela margem para o 2º
   colocado. Se você podar candidatos, mostre que o resultado não piorou.

### Critério de aceite

- Análise completa do deck TIM em **tempo aceitável** (proponha a meta e
  justifique; minutos, não dezenas de minutos).
- Diga **quantos dos 639 alvos passaram a casar** — esse é o ganho de produto que
  motivou tudo. Hoje é 0.
- Deck Natura (`natura_cb/Relatorio_2Q26_Brasil_118_slides.pptx` + `Pedido.xlsx`,
  4 abas) **não pode regredir**: hoje 182 alvos, 3 planos, ~18,5 s, pico 666 MB.
- **137 testes continuam passando**, e acrescente teste para o caminho novo.
- Números antes/depois medidos, não estimados.

### Ferramenta pronta para medir

`scripts/measure_memory_case.py` mede tempo e pico de memória sem alterar os
arquivos de origem:

```
PYTHONPATH=. .venv\Scripts\python.exe scripts\measure_memory_case.py --pptx "<deck>" --xlsx "<planilha>" [--generate]
```

Medições de memória já feitas (2 GB tem ~3x de folga, **não** precisa aumentar):
- TIM: pico **350 MB**
- Natura com geração: pico **666 MB**

Descoberta relevante: memória escala com **trabalho feito**, não com tamanho do
arquivo — o deck de 213 MB gastou metade do de 84 MB. Portanto **não** implemente
aviso de "tamanho combinado máximo": preveria errado.

---

## Também pendente (menor, faça depois do gargalo)

Achados da revisão do seu último commit (`6c5fa80`) — o resto estava bom:
isolamento por squad é real, todas as 34 rotas cobertas, senha compartilhada
desligada em produção, barra de porcentagem real, budget US$20 ativo.

1. **`_job_squad_from_runtime_path` falha aberto** (`web/main.py` ~L191): se o
   `metadata.json` do job existir mas estiver ilegível, retorna `""` e o acesso é
   **liberado**. Deve negar quando o job existe e o squad não pode ser
   determinado.
2. **Duas fórmulas de porcentagem**: uma usa `max(total + 1, 1)`, outra
   `max(total, 1)`. Mesmo conceito, contas diferentes.
3. **Sem alarme de erro.** Existe budget (custo), mas nada avisa se o app começar
   a responder 5xx ou se o health check oscilar. Proponha algo simples
   (CloudWatch + SNS por e-mail).
4. **S3 sem regra de expiração.** Cada execução guarda entrada **e** saída (TIM ≈
   427 MB por execução), com versionamento ligado e nada expirando: cresce sem
   teto. Proponha ciclo de vida.
5. `tim_memory.json` foi salvo sem o campo de pico de memória — o benchmark não
   gravou justamente o número que existia para capturar.
6. **Limite de upload**: `AUTO_PPT_MAX_UPLOAD_MB=250` e o maior deck real tem 213
   MB (85% do limite). Um deck pouco maior é recusado.

## Depois disso, o maior ganho de produto

**Escrever alvos `text`/`shape`.** A descoberta acha 4 tipos (`chart`, `table`,
`text`, `shape`), mas o writer grava só `chart` e `table`
(`engine.py`: `if target.object_type not in {"chart", "table"}`). Deck de 639
objetos tem muito KPI em caixa de texto, hoje simplesmente ignorado.
