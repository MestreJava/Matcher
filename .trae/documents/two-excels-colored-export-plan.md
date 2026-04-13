# Plano: migrar de 1 Excel/2 abas para 2 Excels separados com exportacao colorida

## Resumo
- Objetivo: substituir a entrada atual baseada em `1 arquivo Excel + sheet_t1 + sheet_t2` por `2 arquivos Excel separados`, mantendo compatibilidade com o fluxo de analise, revisao e exportacao.
- Resultado esperado:
  - o GUI passa a aceitar `Excel 1` e `Excel 2` como arquivos distintos;
  - o sistema identifica as abas de cada arquivo e usa por padrao a primeira aba, mas permite troca pelo usuario;
  - o arquivo exportado final passa a ter `4 abas principais`:
    - aba 1: conteudo original do Excel 1, sem colunas extras, com linhas pintadas;
    - aba 2: conteudo original do Excel 2, sem colunas extras, com linhas pintadas;
    - aba 3: resultados de match, com modo de saida `enxuta` ou `tecnica` definido no GUI;
    - aba 4: conciliacao agrupada por nome, mostrando os valores pareados dos dois lados e as sobras por quantidade;
  - o GUI permite escolher cores para:
    - match exato por letra/posicao/tamanho;
    - match aceito que nao seja exato;
    - casos em revisao;
    - sem match;
  - a pintura deve aparecer nas abas 1, 2, 3 e 4;
  - o GUI passa a oferecer uma opcao de comportamento para excesso de repeticoes no Excel 1 quando houver menos ocorrencias equivalentes no Excel 2.

## Current State Analysis
- O app atual usa um unico arquivo em `matching_nomes_gui_v2.py` com:
  - configuracao `input_file`, `sheet_t1`, `sheet_t2`, `header_row_t1`, `header_row_t2`, `name_col_t1`, `name_col_t2`;
  - leitura das duas tabelas por `prepare_input_frames()` usando o mesmo `input_file`;
  - preview de workbook por `collect_workbook_preview()` assumindo um unico Excel com 2 abas;
  - validacao por `validate_config()` assumindo um unico `input_file`;
  - exportacao por `export_analysis_result()` gerando varias abas tecnicas;
  - pintura por `format_output_workbook()` usando buckets fixos (`GREEN_160`, `GREEN_220`, `BLUE_170`, `RED_200`) e aplicando formatacao principalmente nas abas de saida.
- A pintura atual depende de `final_color_bucket`, calculado em `recompute_final_state()` via `determine_color_bucket()`.
- O README ja antecipa evolucao para comparar arquivos inteiros, o que torna esta mudanca coerente com a direcao do projeto.

## Assumptions & Decisions
- Entrada:
  - o usuario vai informar dois arquivos distintos: `input_file_t1` e `input_file_t2`;
  - o sistema deve listar quantas abas existem em cada arquivo e deixar a primeira como default;
  - o usuario ainda podera escolher a aba desejada de cada arquivo no GUI.
- Exportacao:
  - as abas 1 e 2 do arquivo final devem preservar apenas as colunas originais dos arquivos fonte;
  - nessas abas, o sistema nao adicionara colunas auxiliares; o feedback sera por pintura de linha/celula;
  - a aba 3 tera modo configuravel:
    - `enxuta`: colunas essenciais;
    - `tecnica`: colunas detalhadas, proxima da saida atual.
  - a aba 4 sera sempre uma conciliacao agrupada, em duas colunas, com linhas em branco entre grupos.
- Cores:
  - as categorias de cor serao:
    - `EXACT`: match exato de letra/posicao/tamanho;
    - `MATCH`: match aceito nao exato;
    - `REVIEW`: casos de revisao;
    - `NO_MATCH`: sem match;
    - `EXCESS_LEFT`: excedente de quantidade do Excel 1 sem par correspondente no Excel 2, com default `BLUE_160`;
  - o GUI tera cores padrao editaveis e persistiveis.
- Regra de quantidade:
  - sera adicionada opcao no GUI para controlar o comportamento de repeticoes quando a cardinalidade dos nomes divergir;
  - regra solicitada para este plano:
    - se Excel 1 tiver 3 ocorrencias de `JOAO PEDRO` e Excel 2 tiver 1, apenas 1 ocorrencia do Excel 1 sera considerada casada;
    - as outras 2 ocorrencias do Excel 1 ficam sem par por quantidade e devem ser destacadas com a cor de excedente (`EXCESS_LEFT`, default azul 160);
    - a conciliacao da aba 4 deve evidenciar essa sobra.
- Qualidade:
  - a implementacao deve preservar o motor atual de scoring e quota, mudando somente o pipeline de entrada/exportacao e a camada de colorizacao;
  - a mudanca deve ser feita de forma incremental e verificavel para minimizar risco de regressao.

## Proposed Changes

### 1. Reestruturar configuracao de entrada
- Arquivo: `matching_nomes_gui_v2.py`
- Alteracoes:
  - substituir `input_file` por `input_file_t1` e `input_file_t2`;
  - manter `sheet_t1` e `sheet_t2`, mas agora cada um referencia um arquivo distinto;
  - adicionar estado/configuracao para:
    - `output_mode` com valores `enxuta` e `tecnica`;
    - `color_exact`, `color_match`, `color_review`, `color_no_match`;
    - `color_excess_left`;
    - `quantity_resolution_mode` para controlar o tratamento de sobras por quantidade;
  - manter compatibilidade do restante do pipeline usando um `config` normalizado.
- Motivo:
  - o app atual esta rigidamente acoplado a um unico workbook.

### 2. Atualizar GUI para dois arquivos e deteccao de abas
- Arquivo: `matching_nomes_gui_v2.py`
- Alteracoes:
  - trocar a secao de arquivos para:
    - `Planilha Excel 1`
    - `Planilha Excel 2`
    - `Planilha de saida`
  - criar rotina readonly para inspecionar cada arquivo e listar quantidade/nomes das abas;
  - popular `Combobox` de `sheet_t1` e `sheet_t2` automaticamente, selecionando a primeira aba por default;
  - ajustar validacao visual/preview para exibir o resumo dos dois workbooks separadamente;
  - adicionar no GUI:
    - seletor de modo de saida da aba 3;
    - seletores de cor com valores padrao para as 5 categorias;
    - opcao explicita para tratar excesso de repeticoes no Excel 1.
- Como:
  - reaproveitar o padrao atual de `pick_input_file`, `validate_and_preview` e `collect_config_from_vars`;
  - criar helper para ler nomes das abas de cada arquivo e atualizar os comboboxes sem processar os dados completos.

### 3. Refatorar preview e validacao para dois arquivos
- Arquivo: `matching_nomes_gui_v2.py`
- Funcoes afetadas:
  - `collect_workbook_preview()`
  - `validate_config()`
- Alteracoes:
  - validar existencia de `input_file_t1` e `input_file_t2`;
  - validar a aba selecionada em cada arquivo de forma independente;
  - validar a coluna de nome no contexto da aba do arquivo correspondente;
  - atualizar o texto de preview para mostrar:
    - arquivo 1 / abas / aba ativa / amostras;
    - arquivo 2 / abas / aba ativa / amostras.
- Motivo:
  - hoje a validacao assume um unico `pd.ExcelFile(input_file)`.

### 4. Refatorar a leitura dos dados de entrada
- Arquivo: `matching_nomes_gui_v2.py`
- Funcao afetada:
  - `prepare_input_frames()`
- Alteracoes:
  - ler `df1` a partir de `input_file_t1 + sheet_t1`;
  - ler `df2` a partir de `input_file_t2 + sheet_t2`;
  - preservar o restante da normalizacao atual (`source_row_id`, `target_row_id`, nomes normalizados, tokens, prefixos).
- Motivo:
  - manter o core de matching quase intacto reduz risco.

### 5. Introduzir classificacao de pintura baseada em categorias de usuario
- Arquivo: `matching_nomes_gui_v2.py`
- Funcoes afetadas:
  - `determine_color_bucket()`
  - `recompute_final_state()`
- Alteracoes:
  - substituir buckets fixos (`GREEN_160`, `GREEN_220`, `BLUE_170`, `RED_200`) por buckets semanticos:
    - `EXACT`
    - `MATCH`
    - `REVIEW`
    - `NO_MATCH`
    - `EXCESS_LEFT`
  - regra de classificacao:
    - `EXACT`: `final_status == ACEITO` e nome final com 100% por letra/posicao/tamanho;
    - `MATCH`: `final_status == ACEITO` e nao exato;
    - `REVIEW`: `final_status == REVISAR`;
    - `NO_MATCH`: `final_status == SEM_MATCH`.
    - `EXCESS_LEFT`: registro do Excel 1 elegivel no grupo nominal, mas sem par final por falta de quantidade equivalente no Excel 2.
- Motivo:
  - esta regra bate com a preferencia definida pelo usuario: “exato vs resto”.

### 6. Introduzir reconciliacao por quantidade e excedente a esquerda
- Arquivo: `matching_nomes_gui_v2.py`
- Funcoes afetadas:
  - `solve_global_assignment()`
  - `recompute_final_state()`
  - novas helpers de reconciliacao por grupo
- Alteracoes:
  - manter a logica atual de quota e matching global, mas acrescentar uma camada explicita para classificar sobras por quantidade dentro do mesmo nome/grupo;
  - quando houver mais ocorrencias no Excel 1 do que no Excel 2 para um mesmo nome conciliado:
    - apenas a quantidade disponivel no Excel 2 fica casada;
    - o restante do Excel 1 recebe bucket `EXCESS_LEFT`;
  - registrar metadados suficientes para a aba 4:
    - nome/grupo conciliado;
    - ocorrencias lado Excel 1;
    - ocorrencias lado Excel 2;
    - ocorrencias casadas;
    - ocorrencias faltantes por quantidade;
    - lista das sobras do lado esquerdo.
- Motivo:
  - atende diretamente o caso `3 x JOAO PEDRO` no Excel 1 versus `1 x JOAO PEDRO` no Excel 2.

### 7. Redesenhar a exportacao para 4 abas principais
- Arquivo: `matching_nomes_gui_v2.py`
- Funcao afetada:
  - `export_analysis_result()`
- Alteracoes:
  - parar de gerar como saida padrao o conjunto atual de abas tecnicas (`resumo`, `aceitos`, `conflitos`, `candidatos`, etc.);
  - gerar exatamente 4 abas principais:
    - `excel_1_original`
    - `excel_2_original`
    - `resultados_match`
    - `conciliacao_quantidades`
  - construir `excel_1_original` e `excel_2_original` a partir dos dataframes lidos na origem, mantendo apenas colunas originais;
  - montar `resultados_match` em dois formatos:
    - `enxuta`: linhas/ids, nomes, status, score, bucket de cor, match associado;
    - `tecnica`: incluir colunas auxiliares atuais relevantes do processo.
  - montar `conciliacao_quantidades` com 2 colunas principais:
    - coluna 1: valores da coluna selecionada do Excel 1;
    - coluna 2: valores da coluna alvo do Excel 2;
  - ordenar alfabeticamente por grupo de nome;
  - inserir uma linha em branco entre grupos diferentes.
- Observacao:
  - a implementacao deve armazenar ou recomputar referencias suficientes para voltar do `results_df` para os dados originais e colorir corretamente as abas 1 e 2, e para construir a aba 4 com os agrupamentos.

### 8. Pintar abas 1, 2, 3 e 4 sem alterar colunas originais
- Arquivo: `matching_nomes_gui_v2.py`
- Funcao afetada:
  - `format_output_workbook()`
- Alteracoes:
  - aceitar o mapeamento de cores configurado no GUI;
  - aplicar pintura nas 4 abas principais:
    - aba 1:
      - cada linha do Excel 1 recebe cor conforme `final_color_bucket`;
      - os excedentes por quantidade do Excel 1 recebem `EXCESS_LEFT` com default azul 160;
    - aba 2:
      - linhas do Excel 2 que participam de match/review recebem cor coerente com o par correspondente;
      - linhas do Excel 2 nao utilizadas recebem cor `NO_MATCH`;
    - aba 3:
      - pintar cada linha conforme a classificacao `EXACT/MATCH/REVIEW/NO_MATCH/EXCESS_LEFT`;
    - aba 4:
      - pintar os pares e sobras conforme a classificacao do grupo;
      - as linhas em branco entre grupos devem ser preservadas sem pintura.
  - manter cabeçalhos e autosize.
- Ponto de cuidado:
  - para a aba 2, sera necessario definir um criterio consistente quando o mesmo registro aparecer em contexto de revisao ou aceite; o plano de implementacao deve priorizar o status final mais “forte”:
    - `EXACT` > `MATCH` > `REVIEW` > `EXCESS_LEFT` > `NO_MATCH`.

### 9. Adicionar controle de cores no GUI
- Arquivo: `matching_nomes_gui_v2.py`
- Alteracoes:
  - adicionar secao “Cores de pintura” com 5 seletores;
  - defaults sugeridos:
    - `EXACT`: verde forte;
    - `MATCH`: azul;
    - `REVIEW`: amarelo/laranja;
    - `NO_MATCH`: vermelho suave;
    - `EXCESS_LEFT`: azul 160;
  - persistir as cores junto do estado visual salvo.
- Como:
  - usar `tkinter.colorchooser` ou entrada hexadecimal validada;
  - preferencia de implementacao: botoes “Escolher cor” + preview visual + fallback para hex default.

### 10. Adicionar relatorio quantitativo detalhado na aba 3
- Arquivo: `matching_nomes_gui_v2.py`
- Alteracoes:
  - incluir no topo ou bloco inicial da aba `resultados_match` um resumo com:
    - total de linhas do Excel 1;
    - total de linhas do Excel 2;
    - diferenca entre quantidades;
    - total de `ACEITO`, `REVISAR`, `SEM_MATCH` pelo lado do Excel 1;
    - total de `EXCESS_LEFT` no Excel 1;
    - total de linhas utilizadas e nao utilizadas do Excel 2;
    - totais por bucket de cor.
- Motivo:
  - atende o pedido de “quantificar no-matchs/matchs/revisao de cada excel, quantidades totais e diferenca detalhada”.

### 11. Construir a quarta aba de conciliacao agrupada
- Arquivo: `matching_nomes_gui_v2.py`
- Alteracoes:
  - criar dataframe especifico para a aba `conciliacao_quantidades`;
  - regras:
    - duas colunas principais: `Excel 1` e `Excel 2`;
    - cada grupo de nome fica em bloco continuo;
    - ordenar os grupos alfabeticamente;
    - inserir uma linha vazia entre grupos;
    - quando houver diferenca de quantidade:
      - repetir o nome do lado correspondente na quantidade existente;
      - deixar o lado faltante vazio nas linhas excedentes;
      - pintar as sobras do Excel 1 com `EXCESS_LEFT` (default azul 160).
- Exemplo de saida esperado:
  - grupo `JOAO PEDRO`
    - linha 1: `JOAO PEDRO` | `JOAO PEDRO`
    - linha 2: `JOAO PEDRO` | vazio
    - linha 3: `JOAO PEDRO` | vazio
    - linha 4: vazia
- Motivo:
  - atende a solicitacao de conciliacao visual e detalhada por quantidade.

### 12. Preservar uma saida tecnica sem quebrar auditoria
- Arquivo: `matching_nomes_gui_v2.py`
- Alteracoes:
  - no modo `tecnica`, a aba 3 deve conservar as colunas de auditoria necessarias hoje:
    - ids/linhas;
    - scores;
    - gap;
    - flags;
    - metodo final;
    - match final;
    - metricas auxiliares.
- Motivo:
  - o usuario pediu cautela e nao quer perda de capacidade de analise/precisao.

### 13. Ajustar estado salvo e defaults
- Arquivo: `matching_nomes_gui_v2.py`
- Funcoes afetadas:
  - `save_ui_state()`
  - `load_ui_state()`
- Alteracoes:
  - incluir novos campos:
    - `input_file_t1`
    - `input_file_t2`
    - `output_mode`
    - cores configuraveis
    - `quantity_resolution_mode`
  - manter compatibilidade defensiva com estados antigos onde so existe `input_file`.

### 14. Atualizar dependencias apenas se necessario
- Arquivo: `requirements.txt`
- Alteracao planejada:
  - nenhuma nova dependencia obrigatoria para esta feature se for usado `tkinter.colorchooser`;
  - apenas revisar encoding/consistencia do arquivo, se necessario, durante a implementacao.
- Motivo:
  - reduzir risco e evitar ampliar superficie de quebra.

## Implementation Order
1. Ajustar `config` e GUI para dois arquivos e novas opcoes.
2. Refatorar preview/validacao.
3. Refatorar `prepare_input_frames()` para dois arquivos.
4. Preservar o core de matching e adaptar buckets semanticos.
5. Introduzir conciliacao por quantidade e bucket `EXCESS_LEFT`.
6. Reestruturar `export_analysis_result()` para 4 abas principais.
7. Reescrever `format_output_workbook()` para pintura baseada nas cores do usuario.
8. Construir a aba 4 agrupada com separacao por blocos e linhas em branco.
9. Adicionar resumo quantitativo da aba 3.
10. Validar GUI, matching, exportacao e colorizacao.

## Verification Steps
- Validacao estatica:
  - executar `py_compile` no arquivo principal.
- Validacao funcional minima:
  - abrir GUI;
  - selecionar dois arquivos Excel;
  - confirmar populacao automatica das abas com default na primeira;
  - validar preview de ambos os arquivos;
  - rodar analise com dados reais;
  - exportar em modo `enxuta`;
  - exportar em modo `tecnica`;
  - validar a opcao de excesso de quantidade no Excel 1.
- Validacao de integridade do Excel final:
  - confirmar que o arquivo final tem exatamente 4 abas principais;
  - confirmar que as abas 1 e 2 preservam apenas colunas originais;
  - confirmar que as linhas sao pintadas nas abas 1, 2, 3 e 4;
  - confirmar que as cores escolhidas no GUI aparecem corretamente no workbook.
- Validacao de precisao:
  - verificar exemplos de:
    - match exato;
    - match aceito nao exato;
    - revisao;
    - sem match;
    - excesso por quantidade no Excel 1;
  - conferir que a classificacao visual bate com `final_status` e com as metricas de exatidao.
- Validacao da aba 4:
  - confirmar agrupamento alfabetico;
  - confirmar linha em branco entre grupos;
  - confirmar que o caso `3 x JOAO PEDRO` no Excel 1 versus `1 x JOAO PEDRO` no Excel 2 gera 3 linhas no bloco, com 2 sobras pintadas em azul 160.
- Validacao de regressao:
  - confirmar que a logica de quota, revisao manual e exportacao continua funcionando;
  - testar com arquivos de tamanhos diferentes e contabilizar a diferenca na aba 3.

## Risks & Mitigations
- Risco: a logica atual esta acoplada a uma unica entrada.
  - Mitigacao: isolar mudancas em validacao/leitura e preservar o core de matching.
- Risco: a pintura da aba 2 pode ficar ambigua quando um registro participa de multiplos cenarios.
  - Mitigacao: definir prioridade de status final por severidade/forca.
- Risco: a conciliacao por quantidade pode divergir da alocacao global atual.
  - Mitigacao: implementar reconciliacao como camada explicita e testavel, com casos controlados de cardinalidade desigual.
- Risco: perda de informacao ao reduzir as abas de exportacao.
  - Mitigacao: manter modo `tecnica` na terceira aba com dados detalhados.
- Risco: regressao na formatacao do workbook.
  - Mitigacao: validar com arquivos reais e inspecionar workbook gerado apos cada etapa.
