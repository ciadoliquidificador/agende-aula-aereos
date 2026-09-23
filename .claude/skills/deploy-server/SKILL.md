---
name: deploy-server
description: Executa a sequência obrigatória de deploy do server.js (Cia do Liquidificador) na ordem exata do CLAUDE.md — checar sintaxe, commitar, dar push (que dispara o deploy automático no Railway) e confirmar nos logs. Use quando o usuário pedir pra "subir", "fazer deploy" ou "publicar" mudanças no server.js.
disable-model-invocation: true
---

# Deploy do server.js

Execute os passos abaixo **nessa ordem exata, sem pular nenhum**. Se qualquer passo falhar, PARE e reporte o erro ao usuário antes de continuar — nunca prossiga pro próximo passo com um passo anterior quebrado.

## Passo 1 — Checar sintaxe (mandatório, nunca pular)

```bash
node --check server.js
```

Saída silenciosa = OK. Qualquer erro impresso = PARE AQUI. Um erro de sintaxe já quebrou o servidor em produção por esse passo ter sido pulado uma vez — não há exceção a essa regra.

## Passo 2 — Confirmar rotas novas registradas (se houve inserção grande de código)

Se a mudança adicionou rotas novas (`app.get/post/put/delete`), rode:

```bash
grep -n "^app\.\(get\|post\|put\|delete\)(" server.js
```

e confirme visualmente que as rotas novas aparecem na lista, coladas na margem esquerda (nível superior do arquivo, não dentro do corpo de outra função). Já aconteceu de uma rota ficar colada por engano dentro de outra função — sem erro de sintaxe, só nunca registrava (404 permanente).

## Passo 3 — Commit

```bash
git add server.js
git commit -m "<mensagem descritiva da mudança>"
```

Use a mensagem de commit que o usuário forneceu ao invocar esta skill (ou pergunte, se não foi fornecida nenhuma). Nunca use `git add -A` ou `git add .` aqui — sempre `server.js` explicitamente, a menos que o usuário tenha pedido outros arquivos específicos também.

## Passo 4 — Push (isso É o deploy)

```bash
git push origin main
```

O Railway está ligado ao GitHub: todo push na `main` que muda `server.js`, `package.json`, `package-lock.json` ou `railway.json` (ver `watchPatterns` em `railway.json`) dispara o deploy sozinho. Push que só muda outros arquivos (docs, `.claude/`, scripts) é pulado e **não** reinicia o servidor.

**Não rode `railway up`** — geraria um segundo deploy com a pasta local (o hook bloqueia). Todo restart derruba o que está em memória: sessões dos portais, fila de disparos de e-mail, estado da Sala de Ensaio.

## Passo 5 — Acompanhar o deploy do commit

```bash
c=$(git rev-parse --short=7 HEAD); for i in $(seq 1 30); do s=$(railway deployment list --json | jq -r --arg c "$c" '[.[] | select((.meta.commitHash // "") | startswith($c))][0] | "\(.id) \(.status)"'); echo "$s"; case "$s" in *SUCCESS*|*FAILED*|*CRASHED*|*SKIPPED*) break;; esac; sleep 10; done
```

- `SUCCESS` → seguir pro passo 6.
- `FAILED`/`CRASHED` → PARE e mostre `railway logs --deployment <id>` ao usuário.
- `SKIPPED` → o commit não mudou nenhum arquivo dos `watchPatterns`. Se era pra ter deploy, algo está errado: avise o usuário.
- Nada aparece em ~1 min (`null null`) → o webhook do GitHub pode ter falhado. Avise o usuário; ele pode usar "Deploy Latest Commit" no painel do Railway. Não contorne com `railway up`.

## Passo 6 — Confirmar nos logs

```bash
railway logs --deployment <id>
curl -s -o /dev/null -w "%{http_code}\n" https://agende-aula-aereos-production.up.railway.app/health
```

Confirme `Proxy rodando na porta 8080`, sem erro nem stack trace, e health 200. Não diga "deploy concluído" antes disso.

## Ao final

Resuma pro usuário em 2-3 linhas: o que foi commitado, se subiu limpo, e qualquer coisa que precise de atenção manual (ex: automação do Notion pra recriar, variável de ambiente nova pra configurar, etc).
