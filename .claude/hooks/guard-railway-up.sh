#!/bin/bash
# O Railway faz deploy sozinho a cada push na main que muda server.js/package*.json
# (railway.json → watchPatterns). `railway up` sobe a pasta local por fora do GitHub
# e gera um segundo deploy (= segundo restart), então fica bloqueado pro Claude --
# mas só quando o comando é sobre o PROJETO Agende-Aula. Outros projetos irmãos
# (ex: triagem-email) não têm remote no GitHub nem deploy automático nenhum --
# `railway up` é o único jeito deles subirem, e bloquear isso por engano de
# escopo já aconteceu de verdade (set/2026).

cmd=$(jq -r '.tool_input.command // empty' 2>/dev/null)
printf '%s' "$cmd" | grep -qE '(^|[;&|(])[[:space:]]*railway[[:space:]]+up([[:space:]]|$)' || exit 0

# Descobre a pasta-alvo do comando: usa o último "cd <caminho>" que aparece
# antes do "railway up" (padrão desta sessão: "cd /caminho/do/projeto &&
# railway up ..."); sem "cd" nenhum, assume a pasta do projeto atual.
alvo=$(printf '%s' "$cmd" | grep -oE 'cd[[:space:]]+[^;&|]+' | tail -1 | sed -E 's/^cd[[:space:]]+//' | xargs)
alvo="${alvo:-$CLAUDE_PROJECT_DIR}"
alvo="${alvo/#\~/$HOME}"
case "$alvo" in
  /*) ;;
  *) alvo="$CLAUDE_PROJECT_DIR/$alvo" ;;
esac
alvo="${alvo%/}"

if [ "$alvo" != "/Users/fabiospila/Public/Agende-Aula" ]; then
  exit 0  # projeto diferente do Agende-Aula -- essa regra de deploy automático não se aplica
fi

jq -n '{hookSpecificOutput: {hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: "railway up bloqueado: o deploy do Agende-Aula é automático no git push origin main (Railway ligado ao GitHub, railway.json → watchPatterns). railway up geraria um segundo deploy/restart com a pasta local. Se o deploy automático não aparecer, avisar o Fábio em vez de contornar."}}'
