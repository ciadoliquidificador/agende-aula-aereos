#!/bin/bash
# O Railway faz deploy sozinho a cada push na main que muda server.js/package*.json
# (railway.json → watchPatterns). `railway up` sobe a pasta local por fora do GitHub
# e gera um segundo deploy (= segundo restart), então fica bloqueado pro Claude.

cmd=$(jq -r '.tool_input.command // empty' 2>/dev/null)
printf '%s' "$cmd" | grep -qE '(^|[;&|(])[[:space:]]*railway[[:space:]]+up([[:space:]]|$)' || exit 0

jq -n '{hookSpecificOutput: {hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: "railway up bloqueado: o deploy é automático no git push origin main (Railway ligado ao GitHub, railway.json → watchPatterns). railway up geraria um segundo deploy/restart com a pasta local. Se o deploy automático não aparecer, avisar o Fábio em vez de contornar."}}'
