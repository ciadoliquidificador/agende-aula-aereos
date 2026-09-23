#!/bin/bash
# Bloqueia git commit no Agende-Aula se server.js tiver erro de sintaxe.
# Só vale pra pasta do Agende-Aula: sessões abertas aqui também commitam em
# projetos irmãos (ex: triagem-email), que não têm nada a ver com server.js.

cmd=$(jq -r '.tool_input.command // empty' 2>/dev/null)
printf '%s' "$cmd" | grep -qE '(^|[;&|(])[[:space:]]*git[[:space:]]+(-C[[:space:]]+[^[:space:]]+[[:space:]]+)?commit([[:space:]]|$)' || exit 0

# Pasta-alvo: "git -C <pasta>" se houver, senão o último "cd <pasta>", senão o projeto atual.
alvo=$(printf '%s' "$cmd" | grep -oE 'git[[:space:]]+-C[[:space:]]+[^[:space:]]+' | tail -1 | sed -E 's/^git[[:space:]]+-C[[:space:]]+//' | xargs)
[ -z "$alvo" ] && alvo=$(printf '%s' "$cmd" | grep -oE 'cd[[:space:]]+[^;&|]+' | tail -1 | sed -E 's/^cd[[:space:]]+//' | xargs)
alvo="${alvo:-$CLAUDE_PROJECT_DIR}"
alvo="${alvo/#\~/$HOME}"
case "$alvo" in
  /*) ;;
  *) alvo="$CLAUDE_PROJECT_DIR/$alvo" ;;
esac
alvo="${alvo%/}"

[ "$alvo" = "/Users/fabiospila/Public/Agende-Aula" ] || exit 0

cd "$alvo" || exit 0
if err=$(node --check server.js 2>&1); then exit 0; fi
jq -n --arg err "$err" '{hookSpecificOutput: {hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: ("Commit bloqueado: server.js tem erro de sintaxe.\n" + $err)}}'
