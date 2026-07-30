# Preparador de Modelo

O Preparador de Modelo transforma um PowerPoint recorrente em um contrato
reutilizável por squad. Ele funciona sem IA e nunca altera o PPTX original.

## Fluxo

1. `GET /models/{squad}/new`: formulário do PPTX original.
2. `POST /models/{squad}/prepare`: descobre gráficos e tabelas, cria o manifesto
   e uma cópia do PPTX com IDs estáveis nos nomes internos dos objetos.
3. `GET /models/{squad}/studio/{job_id}`: mostra o mapa estrutural e a grade de
   configuração. Alterações são salvas automaticamente e pelo botão `Salvar`.
4. O usuário pode trabalhar na tela ou baixar o XLSX de mapeamento.
5. `POST /models/{squad}/studio/{job_id}/import`: recebe o mapeamento e todos os
   XLSX/ZIP necessários, retém os uploads no rascunho e valida o conjunto.
6. Se houver erro, nada é publicado e a tela lista exatamente o arquivo, a aba
   ou a estrutura que precisa de correção.
7. Se estiver tudo certo, publica uma versão do modelo, salva a memória de
   mapeamento e abre o preview normal com IA desativada.
8. Nas rodadas seguintes, `GET /models/{squad}/{slug}/run` pede somente as
   planilhas atuais, revalida as fontes e abre o preview.

## Colunas da aba `OBJETOS`

| Coluna | Uso |
| --- | --- |
| `ativo` | `1` atualiza o objeto; `0` preserva o objeto intacto. |
| `id_objeto` | ID estável e protegido, por exemplo `S003_T005_CHART`. |
| `slide` | Número do slide. |
| `tipo_objeto` | `chart` ou `table`. |
| `tipo_visual` | Tipo do gráfico ou `tabela`. |
| `titulo_slide` | Melhor título geral detectado no slide. |
| `titulo_detectado` | Título mais próximo do objeto. |
| `confianca_titulo` | `alta`, `media`, `baixa` ou `ausente`. |
| `nome_amigavel` | Nome editável para a pessoa reconhecer o objeto. |
| `linhas_no_ppt` | Categorias/linhas atuais do PowerPoint. |
| `colunas_no_ppt` | Séries/colunas atuais do PowerPoint. |
| `arquivo_xlsx` | Nome exato da fonte. Também aceita `arquivo.xlsx#Aba`. |
| `aba_xlsx` | Nome da aba, quando não foi informado junto ao arquivo. |
| `modo_leitura` | `auto`, `tabela_excel`, `dinamico` ou `exato`. |
| `referencia` | Nome da Tabela do Excel ou range, conforme o modo. |
| `orientacao` | Preferência registrada: `auto`, `manter` ou `transpor`. |
| `formato_valores` | `auto`, `percentual`, `numero` ou `milhares`. |
| `observacao` | Anotação livre do responsável pelo modelo. |
| `status` | Fórmula de conferência na própria planilha. |

Colunas geradas são protegidas. Campos editáveis usam fundo verde e validações
de lista. As abas `_METADADOS` e `_LISTAS` são ocultas e pertencem ao sistema.

## Crescimento mensal sem capturar auxiliares

A regra recomendada é criar uma **Tabela do Excel nomeada** somente sobre os
dados reais e usar:

- `modo_leitura=tabela_excel`
- `referencia=nome_da_tabela`

Quando a tabela ganha novos meses ou atributos, o Excel amplia sua referência e
o sistema lê o novo tamanho. Cálculos auxiliares fora dela não entram. Meses e
categorias novos entram automaticamente no gráfico. Uma série visual inteiramente
nova exige revisão, pois o PowerPoint precisa decidir seu tipo, cor e posição.

Sem uma Tabela nomeada:

- `auto` usa a única Tabela do Excel da aba; se não existir, escolhe o bloco
  principal não vazio;
- `dinamico` usa o range informado como semente e cresce apenas pela região
  contígua;
- `exato` nunca cresce e deve ser usado quando o limite precisa ser rígido.

## Validação

Somente linhas com `ativo=1` exigem fonte. Para cada uma, o sistema:

1. resolve o arquivo sem diferenciar maiúsculas e minúsculas;
2. abre o XLSX e confirma a aba;
3. resolve Tabela do Excel ou range;
4. interpreta a estrutura;
5. executa a mesma normalização do preview;
6. bloqueia a publicação em qualquer erro.

Arquivos extras geram aviso, não erro. Uploads recebidos ficam no rascunho para
que a próxima tentativa possa enviar somente os itens faltantes.

## Persistência e custo

Cada modelo publicado fica em:

```text
squads/{squad}/prepared_models/{slug}/
  model.json
  versions/{version_id}/
    original.pptx
    identified.pptx
    mapping.xlsx
```

`model.json` aponta para a versão atual. Uma versão nova não sobrescreve a
anterior. As fontes XLSX de cada rodada ficam apenas no checkpoint do projeto;
não são copiadas também para o modelo. Isso mantém o fluxo auditável sem criar
chamadas de IA nem duplicação desnecessária no S3.

## Segurança

- O original é lido em memória e nunca gravado de volta no caminho recebido.
- IDs internos são aplicados somente à cópia identificada.
- Modelos e listagens respeitam o isolamento `squad1` a `squad5`.
- O preview recebe `allowed_target_ids`; objetos inativos não entram na análise
  nem na geração.
