# Política de uso da IA

## Padrão

- IA desativada por padrão no preview.
- Nenhuma chamada ao trocar XLSX manualmente.
- Nenhuma chamada ao adicionar slides.
- Revisão somente por ação explícita do usuário, salvo configuração consciente de `AUTO_PPT_AI_AUTO_SOURCE_REVIEW=1`.

## Roteamento inicial

| Operação | Modelo | Reasoning | Motivo |
| --- | --- | --- | --- |
| `source_match` | `gpt-5.6-luna` | `none` | escolha simples entre candidatos já pontuados |
| `transform_diagnostics` | `gpt-5.6-luna` | `none` | auditoria curta e estruturada |
| `slide_matrix_builder` | `gpt-5.6-terra` | `low` | montagem exata de matriz tipada |

`slide_understanding` permanece configurável para compatibilidade, mas o fluxo web
atual consolidou entendimento e montagem em uma única chamada ao matrix builder.

## Controle de dados e custo

- Enviar só datasources relevantes ao target.
- Em override manual, enviar somente o XLSX escolhido.
- Usar dump textual compacto com coordenadas; sem imagem/rasterização.
- Manifesto não repete `preview_rows`.
- Saída da IA contém `value` e `type`; `force_text`, `source_raw` e
  `matrix_preview` são derivados pelo código.
- Rastreabilidade retorna range por fonte, não uma cópia de cada célula.
- `logs/ai_usage.jsonl` não guarda prompt ou resposta. Guarda modelo, reasoning,
  bytes, tokens, tokens em cache, latência e custo estimado.

Preços usados apenas para estimativa local:

- Luna: US$ 1,00/M input, US$ 0,10/M cached input, US$ 6,00/M output.
- Terra: US$ 2,50/M input, US$ 0,25/M cached input, US$ 15,00/M output.

Tokens brutos ficam registrados para recalcular custo se a tabela pública mudar.

## Medição inicial

Caso real `andre/t.xlsx` contra T5 do slide 3:

- dump: 356 bytes;
- manifesto: 819 → 475 bytes;
- payload do target: 773 bytes;
- núcleo repetido em duas chamadas: 3.896 bytes;
- núcleo atual em uma chamada: 1.604 bytes;
- redução: 58,8%.

O override determinístico analisou o deck em 1,545 s com cache frio e 108 ms com
cache quente. A planilha original foi somente lida; medição usou cópia temporária.
