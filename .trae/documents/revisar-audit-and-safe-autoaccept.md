# Plano: Auditoria de REVISAR + Auto-aceite seguro (nome exato + data exata)

## Summary
Fazer uma auditoria focada nos casos `REVISAR` agrupando por “nome repetido” (group key) e por variantes de data, e reduzir a ambiguidade com um ajuste **de baixo risco**: converter automaticamente `REVISAR → ACEITO` apenas quando o nome for exato (exact) e a data (match2) for exatamente igual após normalização `dd/mm/yyyy`, sem sinais de conflito.

Saída do audit: **somente no log/progresso** (sem criar arquivos extras).

## Current State Analysis (grounded)
- Dataset `tests` apresenta `REVISAR` relevantes por grupos repetidos.
- Medição local (script de auditoria):
  - `REVIEW_COUNT = 268`
  - `REVIEW_WITH_EXACT_TARGET_DATE = 125` (existem muitos casos em que há alvo com data exata disponível, mas a linha ainda fica em revisão).
- Causas principais típicas para `REVISAR`:
  - gap baixo (`LOW_GAP`)
  - `STRUCTURE_WARNING`
  - conflito de cota / realocação global
  - nome não é exato (ex.: “MARIA SILVANIA CABRAL” vs “MARIA ILZA DA SILVA”) mesmo com data coincidente → deve continuar `REVISAR`.

## Proposed Changes

### 1) Auditoria focada de `REVISAR` por grupo repetido
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Criar função utilitária (pura) para gerar estatísticas de revisão:
    - total de `REVISAR`
    - top-N grupos por volume (ex.: 20)
    - contagem por “tem alvo com data exata disponível”
    - contagem por motivo/flags (LOW_GAP, STRUCTURE_WARNING, QUOTA_CONFLICT, GLOBAL_REALLOCATED…)
    - exemplos representativos por grupo (até K linhas).
- Como:
  - Rodar após `recompute_final_state()` dentro de `export_analysis_result()` (ou após `analyze_matching()`), apenas quando `progress_callback` existir.
  - Emitir logs concisos (não spammar): 1–2 linhas de resumo + top grupos + exemplos limitados.
- Por quê:
  - Permite “enxergar” onde está a ambiguidade restante antes de ajustar regras.

### 2) Auto-aceite seguro (somente nome exato + data exata)
- Arquivo: `matching_nomes_gui_v2.py`
- O que:
  - Implementar uma regra pós-classificação para promover `REVISAR → ACEITO` **somente** quando:
    - candidato final (ou best/assigned) tem `exact_norm` **ou** `exact_prefix`, e
    - match2 está ativo e **as datas normalizadas** são `dd/mm/yyyy` e são iguais, e
    - não há conflitos críticos (ex.: QUOTA_CONFLICT, FINAL_QUOTA_CONFLICT, GLOBAL_REALLOCATED, needs_length_review).
- Onde aplicar:
  - No ponto em que o sistema define `analysis_status`/`analysis_method` antes de `recompute_final_state()`, ou como um passo explícito logo após `recompute_final_state()` mas antes de exportar.
  - A preferência é aplicar **antes** do `recompute_final_state()` para manter coerência do pipeline.
- Por quê:
  - Esses casos são “quase certamente corretos”, e hoje ficam em `REVISAR` por regras conservadoras (gap/heurísticas), gerando volume alto para revisão manual.

### 3) Garantias (não-regressão)
- Não alterar a lógica para casos onde o nome não é exato.
- Não auto-aceitar se houver qualquer suspeita de conflito (cota, realocação global, etc.).
- Manter as correções já feitas:
  - datas determinísticas `dd/mm/yyyy` (sem flip 06/04 ↔ 04/06)
  - duplicatas com variantes de data no alvo
  - Sheet3 alinhado por “mesmo nome primeiro”

## Assumptions & Decisions
- Decisões do usuário:
  - Auto-aceite: **Sim (seguro)** apenas com nome exato + data exata.
  - Saída do audit: **somente log/console** (sem CSV e sem nova aba no XLSX).
- O “match2” (data) é um critério forte: se a data não bate, não promover para ACEITO.

## Verification Steps
1. Checks estáticos:
  - `py_compile` em `matching_nomes_gui_v2.py`
  - Diagnósticos limpos no editor.
2. Auditoria:
  - Rodar `analyze_matching()` no dataset `tests` e confirmar que o log mostra:
    - total REVISAR
    - top grupos
    - contagens por flags
    - amostras limitadas.
3. Auto-aceite:
  - Confirmar que linhas promovidas atendem rigorosamente:
    - nome exato (exact_norm/exact_prefix)
    - data igual `dd/mm/yyyy`
    - sem conflitos críticos.
4. Regressão:
  - Validar novamente os casos já críticos:
    - `LUCINEIA PEREIRA DA SILVA` (Sheet3 com par em coluna B)
    - `ROSANE BATISTI` (datas 06/04/2026 continuam exatas e não viram 04/06/2026)
  - Confirmar que casos como “MARIA SILVANIA CABRAL” vs “MARIA ILZA DA SILVA” permanecem `REVISAR` (nome não exato).
