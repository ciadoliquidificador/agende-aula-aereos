---
name: deploy-server
description: Executa a sequência obrigatória de deploy do server.js (Cia do Liquidificador) na ordem exata do CLAUDE.md — checar sintaxe, commitar, subir pro GitHub, fazer deploy no Railway e confirmar nos logs. Use quando o usuário pedir pra "subir", "fazer deploy" ou "publicar" mudanças no server.js.
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

## Passo 4 — Push

```bash
git push origin main
```

## Passo 5 — Deploy no Railway

```bash
railway up --detach
```

## Passo 6 — Esperar a propagação

Espere uns 30 segundos antes de checar os logs (o deploy leva um tempo pra subir).

## Passo 7 — Confirmar nos logs

```bash
railway logs
```

Confirme que o serviço subiu limpo, sem erro, sem stack trace. Se aparecer erro, reporte imediatamente ao usuário — não diga "deploy concluído" até confirmar os logs limpos.

## Ao final

Resuma pro usuário em 2-3 linhas: o que foi commitado, se subiu limpo, e qualquer coisa que precise de atenção manual (ex: automação do Notion pra recriar, variável de ambiente nova pra configurar, etc).
