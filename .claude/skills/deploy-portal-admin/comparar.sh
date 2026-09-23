#!/bin/bash
# Compara cada arquivo de ~/Public/portal-admin-deploy com o publicado em
# admin.ciadoliquidificador.com.br (hash SHA-256). Imprime só os diferentes.
# Saída vazia = tudo publicado igual ao local.

ORIGEM="$HOME/Public/portal-admin-deploy"
SITE="https://admin.ciadoliquidificador.com.br"

cd "$ORIGEM" || exit 1
for f in *; do
  [ -f "$f" ] || continue
  [ "$f" = ".DS_Store" ] && continue
  local_hash=$(shasum -a 256 "$f" | cut -d' ' -f1)
  remoto_hash=$(curl -sf "$SITE/$f?nocache=$(date +%s)" | shasum -a 256 | cut -d' ' -f1)
  [ "$local_hash" != "$remoto_hash" ] && echo "$f"
done
exit 0
