# PPTX serverless Open XML POC

Este documento registra o caminho que funcionou para atualizar PowerPoint editavel sem COM/Excel local, visando execucao em Linux/Fargate.

## Conclusao curta

Sim, o caminho e viavel para serverless/container Linux, desde que o writer seja cirurgico:

- nao usar `python-pptx chart.replace_data()` para graficos editaveis;
- nao salvar o workbook embutido inteiro com `openpyxl.save()` como caminho principal;
- nao reempacotar todo o `.pptx`/`.xlsx` com `zipfile` de forma generica;
- alterar somente as partes necessarias e preservar a estrutura ZIP/OPC original;
- atualizar o cache visual do grafico no `ppt/charts/chartX.xml` mantendo formulas, orientacao e relationships coerentes.

O erro "O arquivo vinculado nao esta disponivel" foi resolvido quando preservamos os metadados ZIP/OPC e mudamos apenas o payload necessario.

## Por que o caminho antigo quebrava

1. `chart.replace_data()` reconstrui o workbook do grafico e pode trocar aba/metadados, quebrando o "Editar Dados".
2. `openpyxl.save()` reserializa o `.xlsx` embutido inteiro; em nossos testes removeu/alterou partes e metadados internos.
3. Regravar ZIP com `ZipFile(..., "w")` mudou detalhes como `flag_bits` e `external_attr` das entradas internas do `.xlsx`; o PowerPoint foi sensivel a isso.

No arquivo que funcionou, o embedding original tinha entradas com `flag_bits=6` e `external_attr=0`. O writer generico mudava isso. O writer preservador manteve esses metadados.

## Caminho que funcionou para graficos

### Estrutura

O grafico PowerPoint fica em:

```text
ppt/charts/chartX.xml
ppt/charts/_rels/chartX.xml.rels
ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx
```

O `chartX.xml` tem `c:externalData` com `r:id`. O `.rels` do chart aponta para o workbook embutido com relacionamento do tipo `package`.

### Passos seguros

1. Abrir o `.pptx` como ZIP.
2. Descobrir o `chartX.xml` e o `ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` do target.
3. Abrir o `.xlsx` embutido como ZIP.
4. Atualizar somente `xl/worksheets/sheet1.xml` quando a estrutura da planilha ja e a mesma.
5. Preservar todas as outras entradas do `.xlsx` byte-a-byte.
6. Recriar o `.xlsx` embutido preservando:
   - ordem das entradas;
   - `flag_bits`;
   - metodo de compressao;
   - timestamps;
   - `external_attr`;
   - central directory coerente.
7. Substituir somente o embedding no `.pptx`.
8. Para atualizar o visual antes do usuario abrir "Editar Dados", atualizar tambem `ppt/charts/chartX.xml`.
9. Recriar o `.pptx` preservando todas as partes nao alteradas byte-a-byte.

O POC original esta em:

```text
scripts/pptx_openpyxl_surgery_test.py
```

O caminho real do sistema agora usa:

```text
ppt_automator/openxml_zip.py
ppt_automator/embedded_workbook_writer.py
ppt_automator/engine.py
ppt_automator/ppt_chart_writer.py
ppt_automator/ppt_table_writer.py
```

Arquivos de teste que validaram o caminho:

```text
outputs/enxuto_one_cell_10_to_11_correct_orientation_with_cache.pptx
outputs/enxuto_real_values_correct_orientation_with_cache_plano_crescimento.pptx
outputs/enxuto_real_values_visual_1_decimal_editdata_full_precision.pptx
```

## Orientacao do grafico

O cache visual precisa seguir a orientacao real do grafico, nao a orientacao desejada pelo datasource.

No exemplo `Plano de Crescimento (%)`:

```text
series_rows_categories_columns
```

Ou seja:

```text
A2:A7 = series/legenda: Cristal, Bronze, Prata, Ouro, Diamante, Diamante +
B1:D1 = categorias/eixo: Total, Rede 1, Rede 2
B2:D7 = valores
```

Se inverter isso no cache visual, o PowerPoint abre, mas o grafico fica errado.

## Precisao no "Editar Dados" e formato visual

Para manter valores completos no workbook e mostrar apenas uma casa decimal no grafico:

1. Gravar o valor completo no `.xlsx` embutido.
2. Manter o valor numerico completo no `c:numCache`.
3. Configurar o label visual do grafico:

```xml
<c:numFmt formatCode="0.0" sourceLinked="0"/>
```

Resultado:

- no "Editar Dados": `15.9904534606205`;
- no grafico: `16.0`.

Importante: `number_format="0.0"` nao pode ser tratado como percentual. Somente formatos com `%`, como `0.0%`, devem escalar o valor para representacao percentual.

## Caminho para tabelas PowerPoint / DrawingML

O objeto `8282462966` do slide enxuto nao e grafico nem workbook embutido. Ele e uma tabela DrawingML:

```text
p:graphicFrame
  a:graphic
    a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"
      a:tbl
        a:tr
          a:tc
            a:txBody
              a:p
                a:r
                  a:t
```

O arquivo `andre/t.xlsx` foi interpretado como:

```text
Base
Total  = 50
Natura = 20
Avon   = 30
```

O alvo no PPT:

```text
Base:   [vazio]
Total   [vazio]
Natura  [vazio]
Avon    [vazio]
```

Regra que funcionou:

1. Usar a primeira coluna do PPT como chave.
2. Procurar a mesma chave no XLSX.
3. Escrever o valor na segunda coluna da tabela PPT.
4. Preservar `tcPr`, margens, linhas, fonte e estilo existentes.
5. Se a celula estiver vazia e tiver `a:endParaRPr`, inserir o novo `a:r` antes do `a:endParaRPr`. Inserir o run depois do `endParaRPr` deixa o XML com ordem invalida para o PowerPoint e o texto pode nao aparecer.
6. Para herdar a aparencia da celula vazia, copiar o `a:endParaRPr` como `a:rPr` dentro do novo `a:r`.

Arquivo de teste gerado:

```text
outputs/enxuto_table_8282462966_from_txlsx.pptx
```

Resultado esperado:

```text
Base:   [vazio]
Total   50
Natura  20
Avon    30
```

## Teste combinado: grafico editavel + tabela DrawingML

Tambem validamos um PPT com os dois tipos de alteracao no mesmo slide:

```text
outputs/enxuto_cristal_chart_and_base_table.pptx
```

Neste teste:

- o grafico `Plano de Crescimento (%)` (`shape_name=1130655160`) recebeu valores novos somente na serie `Cristal`;
- o workbook embutido manteve valores numericos completos;
- o cache visual do grafico foi atualizado com a orientacao correta;
- o formato visual dos labels ficou `0.0`;
- a tabela de base (`shape_name=8282462966`) foi preenchida com `andre/t.xlsx`.

Partes alteradas no pacote:

```text
ppt/charts/chart4.xml
ppt/embeddings/Microsoft_Excel_Worksheet3.xlsx
ppt/slides/slide1.xml
```

Partes que devem permanecer intactas:

```text
ppt/charts/_rels/chart4.xml.rels
ppt/slides/_rels/slide1.xml.rels
[Content_Types].xml
```

Validacao interna do teste:

```text
Grafico, serie Cristal:
  Total  = 15.9904534606205
  Rede 1 = 21.598968407479
  Rede 2 = 0

Tabela de base:
  Base:   [vazio]
  Total   50
  Natura  20
  Avon    30
```

## Referencias Open XML

- `c:externalData` representa a fonte externa/embutida de dados do grafico:
  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.externaldata?view=openxml-3.0.1
- Tabelas DrawingML usam `a:tbl`:
  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table?view=openxml-3.0.1
- Celulas de tabela usam `a:tc` e contem `a:txBody`:
  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablecell?view=openxml-3.0.1
- Texto visivel em shapes/celulas fica em `a:txBody`, com paragrafos e runs:
  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textbody?view=openxml-3.0.1

## Implicacoes para Fargate/serverless

Este caminho remove a dependencia de:

- Microsoft PowerPoint instalado;
- Microsoft Excel instalado;
- Windows COM;
- desktop/session interativa.

Portanto, e adequado para Fargate/Linux.

Fargate e melhor que Lambda para este caso se:

- os PPTX forem grandes;
- houver IA por slide/target;
- for necessario renderizar preview;
- o processamento puder passar de limites praticos de Lambda.

Lambda pode ser possivel para jobs pequenos, mas Fargate da mais controle de CPU/memoria/tempo e evita limites apertados de payload/temp storage.

## Criterios minimos antes de migrar o sistema real

- Cobrir graficos com orientacao `series_rows_categories_columns`.
- Cobrir graficos com orientacao `categories_rows_series_columns`.
- Cobrir labels de texto forcado como `Nov/25`, `1Q26`, `001`.
- Cobrir numero com muitas casas decimais.
- Cobrir `numFmt` visual (`0.0`, `0.0%`, etc.).
- Cobrir tabelas DrawingML `a:tbl`.
- Garantir que o writer preservador de ZIP seja usado no caminho final.
- Testar abertura no PowerPoint: visual atualizado antes de "Editar Dados" e workbook editavel depois.
