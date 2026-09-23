---
name: check-routes
description: Confere se rotas novas inseridas em server.js (Cia do Liquidificador) foram registradas corretamente no nível superior do arquivo, e não coladas por engano dentro do corpo de outra função. Use depois de qualquer inserção grande de código em server.js, antes de considerar a mudança pronta pra deploy.
---

# Checar registro de rotas em server.js

Já aconteceu (documentado no CLAUDE.md) de uma rota nova ficar colada por engano dentro do corpo de outra função — isso nunca gera erro de sintaxe, então `node --check` passa limpo, mas a rota nunca é registrada de verdade: qualquer chamada pra ela retorna 404 permanente, silenciosamente, sem nenhum log de erro.

## Como checar

1. Identifique quais rotas foram adicionadas ou modificadas na mudança atual (olhe o `git diff server.js` se não souber).
2. Rode:

```bash
grep -n "^app\.\(get\|post\|put\|delete\)(" server.js
```

3. Para cada rota nova esperada, confirme que ela aparece na lista de saída **colada na margem esquerda** (começando exatamente com `app.` na coluna 1, sem espaço/tab antes). Se uma rota nova não aparecer nessa lista — mesmo que o código dela exista em algum lugar do arquivo — é sinal de que ela está aninhada dentro de outra função por engano.
4. Se alguma rota esperada não aparecer, leia o trecho ao redor de onde ela foi inserida e confirme se ela está dentro de alguma chave `{` de outra função que não foi fechada antes dela, ou se falta uma chave de fechamento `}` da função anterior.

## Relato

- Se todas as rotas novas aparecerem corretamente na lista: diga isso explicitamente ("todas as N rotas novas estão registradas corretamente").
- Se alguma não aparecer: aponte exatamente qual rota, onde ela está no arquivo, e o que precisa ser corrigido (geralmente uma chave de fechamento faltando ou sobrando antes da rota).
