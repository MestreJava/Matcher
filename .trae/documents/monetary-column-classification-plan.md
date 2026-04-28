# Plano: Classificação de Coluna Monetária na Aba 4

## Resumo
Adicionar uma opção na GUI para marcar colunas selecionadas em `Extras aba 4` como monetárias, limitada ao escopo da aba de reconciliação (`conciliacao_quantidades`). Quando essas colunas forem exportadas, o sistema deverá reconhecer tanto valores numéricos simples (`1234,56`) quanto textos já formatados (`R$ 1.234,56`) e gravá-los no Excel como número, aplicando formato monetário `R$`.

## Análise do Estado Atual
- O fluxo relevante está todo em `matching_nomes_gui_v2.py`.
- As colunas adicionais da aba 4 são configuradas pela GUI usando:
  - `tab4_extra_cols_t1`
  - `tab4_extra_cols_t2`
- A configuração é montada na aba `Saída e cores` em `_build_config_tab()`, com os campos criados por `_add_csv_column_field()`.
- Hoje essas colunas são tratadas apenas como texto:
  - `parse_csv_columns()` / `serialize_csv_columns()` manipulam listas CSV simples.
  - `build_grouped_reconciliation_df()` injeta as colunas extras com `str(...)`, em campos `E1:<coluna>` e `E2:<coluna>`.
  - `export_analysis_result()` grava a aba `conciliacao_quantidades` via `to_excel(...)`.
  - `format_output_workbook()` só aplica cabeçalhos, filtros, autofit e preenchimentos por bucket; não há formatação numérica/monetária.
- O app lê planilhas com `dtype=str`, então qualquer valor monetário de origem chega como string no pipeline.
- A prévia de validação (`collect_workbook_preview()`) hoje mostra extras selecionadas, mas não distingue tipo monetário.

## Mudanças Propostas

### 1. Configuração: declarar colunas monetárias da aba 4
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Adicionar duas novas chaves de configuração:
    - `tab4_money_cols_t1`
    - `tab4_money_cols_t2`
  - Reutilizar o mesmo padrão CSV já usado pelas extras.
- Como:
  - Incluir novas `tk.StringVar` em `MatcherApp.vars`.
  - Normalizar em `validate_config()` com `parse_csv_columns()` + `serialize_csv_columns()`.
  - Validar que as colunas monetárias escolhidas existem nas colunas do respectivo Excel.
  - Validar que colunas monetárias sejam subconjunto das colunas exportadas em `tab4_extra_cols_t1/tab4_extra_cols_t2`, ou decidir automaticamente adicioná-las nas extras.
- Decisão:
  - O plano adota comportamento explícito e seguro: se uma coluna for marcada como monetária, ela deve também estar presente nas extras da aba 4; se não estiver, a validação deve acusar erro claro em vez de inferir silenciosamente.

### 2. GUI: opção para classificar coluna como monetária
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Adicionar, na seção `Saída e cores`, um novo campo por lado:
    - `Colunas monetárias aba 4 (Excel 1)`
    - `Colunas monetárias aba 4 (Excel 2)`
- Como:
  - Reaproveitar `_add_csv_column_field()` ou extrair um helper genérico para campos CSV de colunas.
  - Manter a UX consistente com o padrão atual (`Entry` + botões auxiliares).
  - Adicionar tooltip explicando:
    - o efeito só vale para a aba `conciliacao_quantidades`
    - aceita números simples e textos com `R$`
    - a coluna deve estar incluída nas extras exportadas.
- Por quê:
  - É o ponto da GUI onde o usuário já escolhe as colunas extras da aba 4, então é a menor mudança com maior coerência.

### 3. Parsing monetário: converter texto para número Excel
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Criar helper dedicado para parsing monetário, por exemplo:
    - `parse_brl_currency_value(value: Any) -> float | None`
- Como:
  - Aceitar:
    - `1234,56`
    - `1.234,56`
    - `R$ 1.234,56`
    - `1234.56`
    - valores já numéricos
  - Rejeitar textos vazios/inválidos retornando `None`.
  - Preservar sinal negativo se existir.
  - Tratar `.` e `,` com heurística compatível com padrão brasileiro:
    - se houver `,`, tratá-la como decimal e remover `.` de milhar
    - se houver só `.`, aceitar como decimal
    - remover `R$`, espaços e caracteres não numéricos relevantes.
- Por quê:
  - O dado entra como string, mas o Excel só formatará corretamente como moeda se a célula receber número real.

### 4. Exportação: manter metadado de quais colunas são monetárias
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Expandir `build_grouped_reconciliation_df()` para carregar as listas monetárias por lado.
  - Parar de forçar `str(...)` para colunas monetárias ao preencher `rows[-1][f"E1:{col}"]` e `rows[-1][f"E2:{col}"]`.
- Como:
  - Para colunas marcadas monetárias:
    - passar o valor por `parse_brl_currency_value(...)`
    - gravar `float` quando conversão funcionar
    - se não converter, manter string original para não perder informação.
  - Para colunas não monetárias:
    - manter comportamento atual em string.
- Decisão:
  - Não alterar `excel_1_original`, `excel_2_original` nem `resultados_match`.
  - O escopo fica só na aba `conciliacao_quantidades`, conforme decisão do usuário.

### 5. Formatação do workbook: aplicar `R$` nas colunas monetárias
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Em `format_output_workbook()`, detectar as colunas `E1:<col>` e `E2:<col>` configuradas como monetárias na aba `conciliacao_quantidades` e aplicar formato monetário Excel.
- Como:
  - Reaproveitar `_find_header_index(...)` para encontrar os headers de cada coluna monetária visível.
  - Aplicar `number_format` nas células de dados dessas colunas:
    - formato proposto: `R$ #,##0.00`
    - opcionalmente, variante com negativos: `R$ #,##0.00;[Red]-R$ #,##0.00`
  - Só aplicar quando o valor da célula for numérico.
- Por quê:
  - Isso garante visual `R$` no Excel sem transformar tudo em texto.

### 6. Prévia e validação: tornar a configuração visível
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Atualizar `collect_workbook_preview()` para exibir também as colunas monetárias escolhidas.
  - Melhorar mensagens de validação quando uma coluna monetária não estiver nas extras ou não existir no cabeçalho.
- Como:
  - Acrescentar linhas de preview como:
    - `Monetárias aba 4 E1: ...`
    - `Monetárias aba 4 E2: ...`
- Por quê:
  - Ajuda o usuário a revisar o mapeamento antes de processar/exportar.

## Assunções e Decisões
- Escopo funcional:
  - Aplicar monetário somente à aba `conciliacao_quantidades` (aba 4).
- Entrada aceita:
  - Tanto valores simples quanto valores já formatados com `R$`.
- Compatibilidade:
  - Nenhuma mudança no matching em si.
  - Nenhuma mudança nas abas `excel_1_original`, `excel_2_original` e `resultados_match`.
- Persistência:
  - As novas configurações monetárias devem ser salvas e recarregadas junto com o estado visual existente.
- Falha de parsing:
  - Se um valor de coluna marcada como monetária não puder ser convertido, manter o texto original na célula, sem interromper exportação.

## Arquivos Afetados
- `matching_nomes_gui_v2.py`
  - Adição das novas variáveis/configs
  - Validação das colunas monetárias
  - Helper de parsing monetário
  - Alteração da geração da aba `conciliacao_quantidades`
  - Aplicação de `number_format` no workbook
  - Ajustes na GUI e na prévia de validação

## Passos de Implementação
1. Adicionar novas chaves de configuração e variáveis Tk:
   - `tab4_money_cols_t1`
   - `tab4_money_cols_t2`
2. Estender `validate_config()` para normalizar e validar essas colunas.
3. Adicionar os novos campos na GUI em `Saída e cores`.
4. Criar helper de parsing monetário BRL.
5. Ajustar `collect_workbook_preview()` para exibir as seleções monetárias.
6. Ajustar `build_grouped_reconciliation_df()` para converter valores monetários em `float` quando possível.
7. Ajustar `format_output_workbook()` para aplicar formato `R$` nas colunas monetárias da aba `conciliacao_quantidades`.
8. Garantir que `save_ui_state()` / `load_ui_state()` preservem as novas configurações.

## Verificação
- Validação manual:
  - Selecionar uma coluna extra da aba 4 com valores monetários no Excel 1 ou Excel 2.
  - Marcar essa mesma coluna como monetária na GUI.
  - Executar análise e exportação.
  - Confirmar na aba `conciliacao_quantidades`:
    - células numéricas reais no Excel
    - exibição com `R$`
    - sem perda de valores inválidos/textuais.
- Casos de teste esperados:
  - `1234,56` exporta como número com formato `R$`
  - `R$ 1.234,56` exporta como número com formato `R$`
  - valor inválido permanece texto
  - coluna monetária fora das extras gera erro de validação claro
  - configurações monetárias persistem após fechar/reabrir a UI
