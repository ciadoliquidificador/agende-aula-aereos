---
name: deploy-portal-admin
description: Publica mudanças do Portal Admin (~/Public/portal-admin-deploy) no FTP da Locaweb (/public_html/admin/) e confirma que o que está no ar é idêntico ao local. Use quando o usuário pedir pra "subir", "publicar" ou "fazer deploy" do portal admin ou de alguma tela dele (recebimentos.html, disparos.html etc).
disable-model-invocation: true
---

# Deploy do Portal Admin

O Portal Admin é HTML estático em `~/Public/portal-admin-deploy/` (não é repositório git). Sobe por FTP com `deploy-ftp-admin.py`, que fica em `~/Public/Agende-Aula/` e lê as credenciais das variáveis do Railway (`FTP_ADMIN_*`). Nunca imprima nem peça essas credenciais.

Execute nessa ordem. Se algum passo falhar, PARE e reporte.

## Passo 1: descobrir o que mudou

```bash
bash ~/Public/Agende-Aula/.claude/skills/deploy-portal-admin/comparar.sh
```

Lista os arquivos cujo conteúdo local difere do publicado em `admin.ciadoliquidificador.com.br` (comparação por hash). Saída vazia = nada a subir, então diga isso ao usuário e pare.

Mostre a lista ao usuário antes de subir. Se aparecer arquivo que ele não mencionou, pergunte se é pra subir junto (pode ser trabalho em andamento).

## Passo 2: service worker (só se mudou CSS, ícone ou manifest)

O `sw.js` serve `portal-admin.css`, `manifest.json` e os ícones do cache primeiro e só atualiza por trás. Então quem tem o app instalado no celular vê a versão velha no primeiro acesso. HTML não tem esse problema (sempre busca da rede primeiro).

Se a lista do passo 1 incluir `portal-admin.css`, `manifest.json`, `logo192.png` ou `logo512.png`, incremente `CACHE_NAME` em `sw.js` (ex: `'portal-admin-v1'` → `'portal-admin-v2'`) e inclua `sw.js` no upload.

## Passo 3: subir

Rodar a partir de `~/Public/Agende-Aula/`, passando só os arquivos da lista:

```bash
railway run python3 deploy-ftp-admin.py arquivo1.html arquivo2.css
```

Nunca rode sem argumentos (sobe tudo) a menos que o usuário peça.

## Passo 4: confirmar

```bash
bash ~/Public/Agende-Aula/.claude/skills/deploy-portal-admin/comparar.sh
```

Tem que sair vazio. Se algum arquivo ainda aparecer, pode ser cache da Locaweb: espere uns 30s e rode de novo antes de reportar erro.

## Ao final

Diga em 1-2 linhas quais arquivos subiram e que o conteúdo publicado confere com o local. Se mudou CSS/ícone, lembre que no app instalado pode ser preciso fechar e abrir de novo uma vez.
