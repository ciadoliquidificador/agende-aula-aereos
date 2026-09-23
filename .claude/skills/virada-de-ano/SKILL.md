---
name: virada-de-ano
description: Checklist da virada de ano da Cia do Liquidificador. Confere se os bancos de orçamento do ano novo existem no Notion, guia a recriação manual das 5 automações do Notion que chamam o server.js e revisa o código que ainda aponta pra bancos fixos de 2026. Use quando chegar o WhatsApp "Bancos criados automaticamente no Notion", ou em dezembro pra se preparar pro ano seguinte.
disable-model-invocation: true
---

# Virada de ano

Argumento: o ano novo (ex: `/virada-de-ano 2027`). Se não vier, use o ano seguinte ao atual.

## Passo 1: os bancos do ano já existem?

```bash
railway run node .claude/skills/virada-de-ano/checar-bancos.js <ano>
```

Só consulta, nunca cria. Os bancos `<ano> - PROPOSTAS E CONTRATOS` e `APRESENTAÇÕES <ano>` nascem sozinhos na primeira vez que a calculadora de orçamento salva uma data daquele ano (`garantirBancosOrcamentoDoAno` no server.js, clonando o molde de 2026). Se ainda não existem, diga isso ao usuário: o passo 2 só pode ser feito depois que existirem. **Não** crie os bancos chamando `/orcamento/salvar-notion` sem o usuário pedir.

## Passo 2: recriar as 5 automações do Notion (manual, na UI)

Automação do Notion não é acessível por nenhuma API, então isso é sempre trabalho manual do Fábio. Mostre a tabela e acompanhe item por item. Em cada uma: abrir o banco de 2026, copiar a automação equivalente e trocar só o banco. A URL do webhook é a mesma de 2026.

Base: `https://agende-aula-aereos-production.up.railway.app`

| # | Banco do ano novo | Dispara quando | Chama (POST) |
|---|---|---|---|
| 1 | PROPOSTAS E CONTRATOS | Status vira "Aprovado - Aguardando contrato" | `/webhook-proposta-aprovada` |
| 2 | PROPOSTAS E CONTRATOS | Cachê vira "PAGO" | `/webhook-cache-pago` |
| 3 | APRESENTAÇÕES | Apresentação criada/editada | `/webhook-apresentacao-notion` |
| 4 | APRESENTAÇÕES | ELENCO / Produção Liqui / Técnico de Som / Técnico de Luz preenchidos | `/webhook-apresentacao-escalacao` |
| 5 | APRESENTAÇÕES | Local Saída / Horário de Saída preenchidos | `/webhook-apresentacao-saida` |

Pra confirmar que uma automação funciona: editar uma página de teste no banco novo e procurar a rota correspondente em `railway logs` (ex: `[webhook-proposta-aprovada]`).

## Passo 3: código que ainda aponta pra 2026

Rode `grep -n "2026" server.js` e revise com o usuário. Pontos conhecidos (set/2026):

- **`APRESENTACOES_2026_DB` fixo** em `/portal-artista/apresentacoes` (anos ≠ 2026 retornam lista vazia com aviso, decisão consciente do Fábio) e em `/admin/sincronizar-apresentacoes-pendentes` (backfill manual). O `/webhook-cache-pago` já acha o banco do ano pelo `parent` da proposta (corrigido em set/2026); se aparecer `banco da proposta (...) não reconhecido` no log, investigar.
- **`PRESENCAS_2026_DB`** em "Minha Presença" do Portal Aluna. Perguntar ao Fábio se em 2027 continua o mesmo banco ou nasce um "Presenças 2027"; se nascer, o portal precisa ler dos dois.
- **Textos com 2026**: `'Turma Única 2026'` (Meditação), textos da Residência Artística, `VERSAO_TEXTO_CONTRATO`. Revisar se fazem sentido no ano novo.

Qualquer mudança no server.js segue o fluxo normal (`/deploy-server`).

## Ao final

Resuma: bancos existem ou não, quais das 5 automações o Fábio confirmou, e o que ficou pendente no código.
