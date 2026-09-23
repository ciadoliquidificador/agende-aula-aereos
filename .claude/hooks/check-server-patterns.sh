#!/bin/bash
# Padrões que já quebraram produção (CLAUDE.md) e que node --check não pega.
# Olha só as linhas adicionadas (git diff) de server.js, ignorando comentários.

f=$(jq -r '.tool_input.file_path // .tool_response.filePath // empty')
case "$f" in */server.js|server.js) ;; *) exit 0 ;; esac
cd "${CLAUDE_PROJECT_DIR:-/Users/fabiospila/Public/Agende-Aula}" || exit 0

added=$(git diff -U0 -- server.js | grep -E '^\+[^+]' | grep -vE '^\+[[:space:]]*(//|\*)')
[ -z "$added" ] && exit 0

problems=""
check() {
  local hits
  hits=$(printf '%s\n' "$added" | grep -E -- "$2" | head -3 | cut -c1-160)
  [ -n "$hits" ] && problems+="• $1"$'\n'"$hits"$'\n\n'
}

check "Regra 5: scheduledAt do Digisac não funciona, usar a fila Notion (📤 Fila de Mensagens Agendadas)" 'scheduledAt'
check "Regra 6: nunca usar data[0].id como fallback de contactId" 'data\[0\]\.id'
check "Regra 8: retornar timestamp absoluto (proximo.getTime()), nunca duração" 'return[[:space:]]+\(?[A-Za-z_.]+[[:space:]]*-[[:space:]]*agora'
check "Regra 4: rota com indentação, provavelmente colada dentro de outra função (nunca registra, 404)" '^\+[[:space:]]+app\.(get|post|put|patch|delete|all)\('

[ -z "$problems" ] && exit 0

jq -n --arg p "$problems" '{
  decision: "block",
  reason: ("Padrões proibidos pelo CLAUDE.md nas linhas alteradas de server.js:\n\n" + $p + "Corrigir antes de commitar."),
  systemMessage: ("⚠️ server.js: padrão proibido pelo CLAUDE.md\n\n" + $p)
}'
