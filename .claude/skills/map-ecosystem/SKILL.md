---
name: map-ecosystem
description: Atualiza o mapa do ecossistema inteiro da Cia do Liquidificador (~/Public/CODEBASE_MAP.md) — todos os ~30 projetos irmãos do Agende-Aula (apps de agendamento, Portal Admin, triagem de e-mail, portais self-service, etc). Diferente do /cartographer padrão, que só mapeia a pasta do projeto atual. Use quando o usuário pedir explicitamente pra atualizar/regerar o mapa do ecossistema, ou avisar que mudou algo em algum dos projetos irmãos em ~/Public.
disable-model-invocation: true
---

# Atualizar o mapa do ecossistema (~/Public)

Isso é uma extensão do `/cartographer` (bundled skill) pra cobrir `~/Public/` inteiro, não só a pasta do projeto atual (`Agende-Aula`). O `/cartographer` puro sempre mapeia só o diretório de trabalho atual — essa skill existe justamente pra ir além disso quando o pedido é sobre o ecossistema todo.

## Passo 1 — Descobrir o escopo atual

```bash
find /Users/fabiospila/Public -maxdepth 1 -type d | sort
```

Compare com a tabela de "Inventário de Projetos" em `~/Public/CODEBASE_MAP.md` (se o arquivo existir). Pastas novas que não aparecem na tabela = projetos novos, precisam de mapeamento do zero. Pastas que sumiram = remover da tabela.

## Passo 2 — Descobrir o que mudou desde o último mapeamento

Leia o `last_mapped` do frontmatter de `~/Public/CODEBASE_MAP.md` (se não existir, é a primeira vez — pule pro Passo 3 mapeando tudo).

Para cada projeto já mapeado na tabela, cheque se algo mudou desde `last_mapped`:

```bash
# Se o projeto tem .git:
cd "<pasta-do-projeto>" && git log --oneline --since="<last_mapped>"

# Se não tem .git (várias pastas estáticas não têm remote nem repo):
find "<pasta-do-projeto>" -newer /tmp/marker-com-timestamp-do-last_mapped -type f -not -path "*/node_modules/*"
```

(Pra comparar por timestamp sem `.git`, crie um arquivo marcador vazio com `touch -d "<last_mapped>" /tmp/marker` — ou, mais simples, use `find ... -mtime -Nd` estimando quantos dias se passaram desde `last_mapped`.)

Projetos sem nenhuma mudança: **não remapeie** — mantenha a seção existente do `~/Public/CODEBASE_MAP.md` como está, só ajuste o `last_mapped` geral no fim.

## Passo 3 — Rodar o scanner nos projetos que mudaram (ou em tudo, se for a primeira vez)

Localize o script do cartographer dinamicamente (o caminho tem a versão do plugin, que pode mudar):

```bash
find /Users/fabiospila/.claude/plugins/cache/cartographer-marketplace -name "scan-codebase.py"
```

Rode contra `~/Public` (ou só contra as pastas que mudaram, se for atualização parcial):

```bash
cd /Users/fabiospila/Public && python3 <caminho-do-script> . --format json > /tmp/scan_public.json
```

Se o `tiktoken` não estiver instalado, `pip3 install --quiet --user tiktoken` primeiro (mesma dependência do `/cartographer` normal).

## Passo 4 — Dividir o trabalho em subagentes (Explore, model sonnet, paralelo, background)

**Nunca leia os arquivos você mesmo — sempre delegue pra subagentes**, do mesmo jeito que o `/cartographer` normal faz, só que agrupando por PROJETO (pasta de primeiro nível dentro de `~/Public`), não por linha de arquivo:

- Orçamento de ~130k tokens por subagente (margem segura sob o limite de 200k do Sonnet).
- Projetos grandes (backend `Agende-Aula`, `triagem-email`, `portal-admin-deploy`) ficam sozinhos num subagente cada.
- Projetos médios/pequenos (os 7 apps de agendamento, os ~15 portais/páginas estáticas menores) são agrupados em 2-3 subagentes, mantendo pastas relacionadas juntas quando possível (ex: todos os apps React de booking num grupo).
- Cada subagente recebe: os caminhos exatos das pastas que ele deve ler, contexto de negócio (Cia do Liquidificador, escola de artes cênicas, backend único em `Agende-Aula/server.js` que todo o resto chama via REST), e o formato de relatório esperado (propósito, arquivos-chave, integração com o backend, padrões notáveis, gotchas — mesmo formato usado no `~/Public/CODEBASE_MAP.md` já existente, pra manter consistência entre atualizações).
- **Não remapeie o backend `Agende-Aula` em detalhe aqui** — esse já tem seu próprio mapa dedicado em `Agende-Aula/docs/CODEBASE_MAP.md`, mantido pelo `/cartographer` padrão rodado ali dentro. Se o backend mudou desde o último `~/Public/CODEBASE_MAP.md`, sugira ao usuário rodar `/cartographer` normal dentro do projeto Agende-Aula pra atualizar aquele mapa específico, e só reflita um resumo curto aqui.

## Passo 5 — Sintetizar

Junte os relatórios dos subagentes num `~/Public/CODEBASE_MAP.md` atualizado:
- Preserve a estrutura existente (Visão Geral / Inventário de Projetos / uma seção por projeto ou grupo / Convenções cross-cutting / Gotchas / Guia de Navegação) — não reinvente o formato a cada atualização.
- Atualize só as seções dos projetos que mudaram; mantenha as demais como estavam.
- Atualize a tabela de Inventário (adicionar projetos novos, remover os que sumiram, atualizar contagem de tokens/arquivos).
- Pegue o timestamp real antes de escrever: `date -u +"%Y-%m-%dT%H:%M:%SZ"` — nunca estime a data.
- Atualize o `last_mapped` do frontmatter.

## Passo 6 — Avisar o usuário

Resuma o que mudou desde o último mapeamento (projetos novos, projetos alterados, achados/gotchas novos) — não repita o mapa inteiro no chat, só o que há de novo. Se algo parecer um problema real (senha exposta, código morto, link quebrado, etc), destaque isso separadamente, como já é hábito nesse projeto.
