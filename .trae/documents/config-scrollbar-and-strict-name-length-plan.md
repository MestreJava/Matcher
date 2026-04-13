# Plano: Scroll na Configuração + limite real de caracteres no Match 1

## Resumo
- Objetivo:
  - adicionar barra de rolagem vertical na tela `Configuração`;
  - corrigir o comportamento do campo `Tamanho do prefixo` para que o limite de caracteres seja efetivamente aplicado no Match 1 (coluna principal de nome).
- Critério de sucesso:
  - a aba `Configuração` permite rolagem suave do conteúdo completo;
  - o limite padrão de caracteres do Match 1 fica em `30`;
  - nomes do Match 1 são truncados para esse limite antes de entrar na lógica de comparação/ranqueamento;
  - o comportamento fica claro no GUI (rótulo/tooltip) para evitar confusão.

## Current State Analysis
- A aba `Configuração` hoje é montada diretamente em `self.tab_config` sem `Canvas + Scrollbar` em [_build_config_tab](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L2279), então em resoluções menores parte das opções pode “sumir”.
- O campo `Tamanho do prefixo` (`max_external_chars`) já existe, porém é usado apenas como:
  - prefixo de indexação em `key_prefix` ([prepare_input_frames](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L697));
  - parcela de pontuação `score_prefix` ([score_candidate](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L221)).
- A comparação principal ainda usa `nome_t1_norm` / `nome_t2_norm` completos, por isso nomes com mais de 30 caracteres continuam participando integralmente em outros sinais (token/order/aligned), causando a percepção de que “não respeita 30”.

## Assumptions & Decisions
- Decisão do usuário:
  - manter padrão `max caracteres = 30`;
  - aplicar limite específico também na coluna principal de match (Match 1).
- Decisão de implementação:
  - o limite será aplicado simetricamente para Match 1 dos dois lados (Excel 1 e Excel 2), para preservar coerência.
  - o campo existente `max_external_chars` será reinterpretado como limite efetivo do Match 1 (não só auxiliar).
  - sem mudança no fluxo de Match 2 além de manter o comportamento já implementado.

## Proposed Changes

### 1. Estruturar `Configuração` com rolagem vertical
- Arquivo: `matching_nomes_gui_v2.py`
- Área: GUI (`MatcherApp._build_ui` e `_build_config_tab`)
- Mudanças:
  - criar container rolável para a aba `Configuração` usando:
    - `tk.Canvas` + `ttk.Scrollbar(orient="vertical")`;
    - frame interno para os blocos de configuração;
    - binding de `<Configure>` para atualizar `scrollregion`;
    - suporte de roda do mouse enquanto a aba estiver ativa.
  - mover os blocos já existentes (`mode_frame`, `files_frame`, `workbook_frame`, etc.) para o frame interno rolável.
- Motivo:
  - acomodar o número crescente de opções sem perda de usabilidade.

### 2. Aplicar limite real no Match 1 antes de comparar
- Arquivo: `matching_nomes_gui_v2.py`
- Área: preparação de dados e scoring
- Mudanças:
  - em `prepare_input_frames()`:
    - manter `nome_t1_norm`/`nome_t2_norm` originais para auditoria;
    - criar campos “efetivos” truncados para matching, por exemplo:
      - `nome_t1_match_norm = nome_t1_norm[:max_external_chars]`
      - `nome_t2_match_norm = nome_t2_norm[:max_external_chars]`
  - em `analyze_matching()`:
    - usar `*_match_norm` em:
      - seleção de pool (`choose_candidate_pool`);
      - `score_candidate`;
      - marcadores de igualdade/estrutura usados no ranking.
  - preservar colunas originais para exibição/export.
- Motivo:
  - garantir que o limite de 30 caracteres seja efetivo no Match 1.

### 3. Ajustar indexação de catálogo para respeitar o limite
- Arquivo: `matching_nomes_gui_v2.py`
- Área: `build_target_catalog()` / índices auxiliares
- Mudanças:
  - basear `first_token`, `last_token`, `key_prefix` e demais índices de busca no campo truncado de match, para manter consistência entre busca e pontuação.
  - manter referência ao nome original completo para saída final.
- Motivo:
  - evitar inconsistência entre fase de candidato e fase de score.

### 4. Clarificar semântica no GUI
- Arquivo: `matching_nomes_gui_v2.py`
- Área: `_build_config_tab()`, preview e mensagens
- Mudanças:
  - renomear texto/tooltip de `Tamanho do prefixo` para algo explícito:
    - ex.: `Limite de caracteres (Match 1)`.
  - no `collect_workbook_preview()`, adicionar linha confirmando:
    - “Match 1 será truncado em N caracteres”.
- Motivo:
  - reduzir ambiguidade operacional.

### 5. Compatibilidade de estado salvo
- Arquivo: `matching_nomes_gui_v2.py`
- Área: `save_ui_state()` / `load_ui_state()`
- Mudanças:
  - manter compatibilidade com estados antigos;
  - garantir que valor default continue `30` quando ausente.

## Verification Steps
- Verificação estática:
  - `py_compile matching_nomes_gui_v2.py`;
  - `GetDiagnostics` sem erros.
- Verificação GUI:
  - abrir app e confirmar scrollbar funcional na aba `Configuração`;
  - testar rolagem por scroll do mouse e barra lateral.
- Verificação funcional do limite:
  - caso com nomes > 30 caracteres (incluindo `NOAH VINICIUS DE CARVALHO PEDROSO`);
  - confirmar, via logs/colunas auxiliares/resultado, que Match 1 usa somente os primeiros 30 caracteres.
- Regressão:
  - validar que Match 2 e aba 4 continuam operando;
  - validar exportação com 4 abas sem quebra de pintura.

## Risks & Mitigations
- Risco: endurecer limite pode alterar resultados anteriores.
  - Mitigação: aplicar de forma simétrica nos dois lados e validar com casos reais/sintéticos antes de liberar.
- Risco: scroll da aba pode conflitar com widgets internos.
  - Mitigação: encapsular scrolling no container da aba e limitar bindings ao contexto da aba `Configuração`.
- Risco: confusão por mudança semântica do campo antigo.
  - Mitigação: atualizar rótulo, tooltip e preview explicitando “limite efetivo”.
