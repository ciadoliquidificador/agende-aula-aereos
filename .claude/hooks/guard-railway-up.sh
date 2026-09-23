#!/bin/bash
# `railway up` sobe a pasta local, não o GitHub. Barra o deploy se o código
# que vai subir não for exatamente o que está commitado e enviado pro origin/main.

cd "${CLAUDE_PROJECT_DIR:-/Users/fabiospila/Public/Agende-Aula}" || exit 0

negar() {
  jq -n --arg r "$1" '{hookSpecificOutput: {hookEventName: "PreToolUse", permissionDecision: "deny", permissionDecisionReason: ("Deploy bloqueado: " + $r + "\nOrdem obrigatória (CLAUDE.md): node --check → git add → git commit → git push origin main → railway up, cada um em comando separado.")}}'
  exit 0
}

branch=$(git rev-parse --abbrev-ref HEAD 2>/dev/null)
[ "$branch" = "main" ] || negar "branch atual é '$branch', não main."

sujos=$(git diff --name-only HEAD -- '*.js' package.json package-lock.json 2>/dev/null)
[ -n "$sujos" ] && negar "arquivos com mudança não commitada: $(echo $sujos)."

pendentes=$(git rev-list --count origin/main..HEAD 2>/dev/null || echo 0)
[ "$pendentes" -gt 0 ] && negar "$pendentes commit(s) ainda não enviados pro GitHub (falta git push origin main)."

exit 0
