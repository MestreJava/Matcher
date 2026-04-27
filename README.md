# MatcherV2

Programa para realizar matching entre duas listas por similaridade (com limiares configuráveis), revisão manual e exportação em Excel.

## Configuração de cores por nome (novo)

Além de códigos hexadecimais, o app agora aceita nomes de cor legíveis na configuração:

- `Verde (Exato)`
- `Azul (Match aceito)`
- `Laranja (Revisão)`
- `Vermelho (Sem match)`
- `Azul claro (Excedente Excel 1)`

Comportamento:

- Ao carregar/salvar estado, nomes de cor continuam compatíveis com códigos já existentes.
- Na validação interna, os nomes são normalizados para fills ARGB do Excel (ex.: `FF4CAF50`).
- Se vier valor inválido, o sistema aplica fallback seguro para a cor padrão daquele bucket.

## Mudanças de comportamento relevantes

- Parse booleano robusto para evitar ambiguidades (`"False"` deixa de ser interpretado como verdadeiro).
- Controle de reaproveitamento de nomes do Excel 2 respeita limite configurado (`max_matches_per_t2_name`) quando habilitado.
- Durante exportação, ações de mutação da revisão manual ficam bloqueadas para evitar corrida de estado.
- Erros de metadados de planilha (abas/colunas/cabeçalho) são retornados com mensagem acionável.

## Smoke check local (Task 5)

Executar com o ambiente virtual do projeto:

```powershell
.\.venv\Scripts\python.exe -m py_compile .\matching_nomes_gui_v2.py
.\.venv\Scripts\python.exe -c "import matching_nomes_gui_v2 as m; print(m.APP_VERSION)"
```

Exemplo de amostra usada nos testes locais:

- Arquivo: `dados.xlsx`
- Aba Excel 1: `Tabela1` (coluna nome `C`)
- Aba Excel 2: `Tabela2` (coluna nome `E`)

## Futuras adições

- Comparar 2 arquivos inteiros, escolhendo colunas que entram no parâmetro de match (datas, códigos etc.).
- Match por posição de caractere além da similaridade.
- Melhorar cenários de limite de quantidade quando há nomes repetidos entre as listas.

