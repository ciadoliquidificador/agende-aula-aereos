# Cia do Liquidificador — Contexto do Projeto

> Este arquivo é lido automaticamente no início de toda sessão do Claude Code.
> Mantenha atualizado conforme o projeto evolui — é a memória de longo prazo do projeto.

## Sobre o negócio

Cia do Liquidificador é uma escola de artes cênicas operando como **Liquidificador Produções Artísticas / Cristiane Socci Leonel - ME** (CNPJ: 28.398.119/0001-89), em Rua Dr. Carvalho de Mendonça, 67 — Campos Elíseos, São Paulo/SP. **Fábio Spila** administra toda a infraestrutura técnica.

**Ano atual: 2026.** Sempre usar 2026 em datas, contratos, nomes de bancos e qualquer documento gerado.

## Stack

- **Frontend:** React apps hospedados na Locaweb via FTP, tema visual creme+vinho (Playfair Display + Inter)
- **Backend:** Um único proxy Node.js/Express compartilhado no Railway (`~/Public/Agende-Aula/server.js`, 13.000+ linhas — mapa completo de seções/rotas/gotchas em [docs/CODEBASE_MAP.md](docs/CODEBASE_MAP.md))
- **Ecossistema completo:** ~30 projetos irmãos vivem em `~/Public/` (os 7 apps de agendamento, o Portal Admin, a triagem de e-mail, portais self-service, etc.) — mapa geral de tudo em [../CODEBASE_MAP.md](../CODEBASE_MAP.md)
- **Deploy backend:** Railway, projeto em `~/Public/Agende-Aula/`
- **Database:** Notion (múltiplos bancos, workspace "Produção Liquidificador")
- **WhatsApp:** Digisac API
- **Calendários:** Google Calendar (service account) + Microsoft OneDrive (Azure App Registration)
- **GitHub:** `https://github.com/ciadoliquidificador/agende-aula-aereos`

## Regras obrigatórias — NUNCA pular

1. **`node --check server.js` é mandatório antes de qualquer `git add`.** Um erro de sintaxe já quebrou o servidor em produção por pular essa etapa.
2. **Sequência de deploy, sempre nessa ordem:**
   ```
   node --check server.js  (silencioso = OK, não prosseguir se der erro)
   git add server.js
   git commit -m "..."
   git push origin main
   railway up --detach
   (esperar ~30s)
   railway logs   (confirmar que subiu sem erro)
   ```
3. **`node server.js` nunca é rodado localmente sozinho** — sempre `railway run node server.js` pra injetar as env vars do Railway, se precisar testar local.
4. Depois de qualquer inserção de bloco grande de código no server.js, rodar `grep -n` pras rotas novas pra confirmar que foram registradas corretamente (já aconteceu de rotas ficarem coladas dentro do corpo de outra função por engano — nunca dava erro de sintaxe, só nunca registrava as rotas, 404 permanente).
5. **Digisac `scheduledAt` não funciona** — confirmado via teste controlado (mensagem "agendada" pra daqui 3min e 20min chegou na hora, imediatamente). Usar sempre a fila própria Notion-backed (banco "📤 Fila de Mensagens Agendadas", polling a cada 60s).
6. **Digisac contactId:** nunca usar fallback `data[0].id` — sempre buscar match exato ou criar novo contato.
7. **`calcularProximoHorarioComercial()`** (e funções de checagem de horário comercial) devem ser `async` e reconhecer feriados municipais/estaduais de SP (9 de julho, 25 de janeiro) + feriados móveis calculados a partir da Páscoa.
8. **Validação de business hours WhatsApp:** usar `return proximo.getTime()` (timestamp absoluto), NUNCA `return proximo - agora` (duração) — esse bug já se repetiu mais de uma vez em apps diferentes (Aéreos, Infantil, Acro).
9. **DDD:** validar `numLimpo.length < 11` antes de mandar qualquer WhatsApp.
10. **Horários passados:** checar `horaAgora >= horaAula` quando o cursor de agendamento é hoje.
11. Nomes de constantes tipo `WHATSAPP_FABIO` — nunca referenciar antes da declaração (já causou crash silencioso 1x em produção).
12. **Ação idempotente disparada por múltiplos webhooks/automações da mesma página do Notion (ex: Apresentações tem 3 gatilhos independentes) precisa de trava em memória por pageId**, não só um checkbox de controle — dois gatilhos quase simultâneos leem o checkbox como `false` antes de qualquer um marcar como `true`, e cada um executa a ação de novo. Confirmado ao vivo (set/2026): lembretes de WhatsApp do produtor duplicaram 3x numa única rajada antes da trava (`Set` de pageIds "em processamento", checado/setado de forma síncrona antes de qualquer `await`). Só funciona porque o Railway roda 1 réplica — se um dia rodar múltiplas réplicas, precisa de lock externo (Notion mesmo, ou Redis).

## Notion — regras e IDs importantes

- **PAT tokens (`ntn_`) são escopados por workspace**, não dão acesso automático a tudo. Bancos herdados via página-pai ("Added via...") são pouco confiáveis via API — só "Selected manually" funciona de forma consistente. Às vezes é preciso um token novo por banco.
- **`ALTER COLUMN ... RENAME`** não é suportado — workaround é `ADD COLUMN "Nome Novo" TYPE` e deixar a coluna antiga sem uso.
- **`ALTER COLUMN SET SELECT(...)`** precisa incluir TODAS as opções já existentes, senão elas somem.
- **Notion Forms:** perguntas e descrições de cabeçalho não podem ser setadas via API — sempre manual na UI do Notion.
- **Money fields:** sempre `number_format: real` (nunca dollar).
- **Se um `POST /v1/pages` der 400 "X is not a property that exists" mesmo o código usando o nome certo**, suspeitar de corrupção no NOME da propriedade em si — já aconteceu (set/2026, banco `📋 Lista de Interesse`) da propriedade título ter ficado literalmente nomeada `Título` (barra invertida e tudo, não o "í" de verdade), provavelmente de algum script que usou a string já com `JSON.stringify` aplicado como nome ao criar a coluna. Sempre conferir `GET /v1/databases/{id}` e comparar bytes do nome antes de assumir que é bug no código — resolve com `PATCH /v1/databases/{id}` renomeando a propriedade pelo nome (quebrado) atual.
- Bancos principais:
  - `Alunas` — `aee12f7f-8cb9-4ee2-80ba-1bcb06d9eda0` (data source `41bb69c4-2d18-4c81-9e43-e60c5f4033f6`)
  - `Presenças 2026` — data source `8365a940-b386-401b-bedb-d26dfff2415e`
  - `🔄 Reposições Solicitadas` — `dde8519e6e0f4157b2bb56b545e2ef84` (data source `53acaa29-45c5-4ace-99b5-8d1149a96e9d`)
  - `📄 Contratos — Professores` — `f1b934a8100143019ac7f5f877c405f1`
  - `📎 Aditivos — Contratos Professores` — `9b35319a73854e318cd9efc6497bfb0e`
  - `📢 Mural de Avisos` — `c45786e213ff463f8558054b2f787a69`
  - `👥 Professores — Cadastro` — `728021ad4c58466db1dd5ab112ada252`
  - `Trabalhos` (catálogo de espetáculos, sem ano) — `4589d769656b41149e9bf6300b30d886`
  - `Integrantes` (catálogo de elenco/técnicos, sem ano) — `e1047585-3dd2-4bda-9896-1a4caeeea284`

## Calculadora de orçamento — bancos anuais criados automaticamente (set/2026)

A calculadora de orçamento (`calculadora_orcamento_v11_1.html`, hospedada fora do FTP — o Fábio abre local) fala com o server.js (`/orcamento/*`). Cada ano tem seu par de bancos no Notion: `"<ano> - PROPOSTAS E CONTRATOS"` (orçamentos/contratos, uma página por data/apresentação) e `"APRESENTAÇÕES <ano>"` (criada automaticamente quando uma Proposta é aprovada, via `/webhook-proposta-aprovada`). Ambos relacionam com os bancos globais `Trabalhos`/`Integrantes` acima (não são por ano).

- **2026 é o molde** (`2c6c45031f73804f8f90e6e7439d7e1c` / `2b9c45031f7380828d34f47353b066e7`) — nunca renomear/apagar.
- **Anos seguintes (2027, 2028, ...) são criados sob demanda** pelo próprio server.js (`garantirBancosOrcamentoDoAno`) na primeira vez que aparece uma data daquele ano em `/orcamento/salvar-notion` — clona o schema completo do molde 2026 (propriedades, opções de select/multi_select, fórmula, rollups, relações). Zero trabalho manual de banco na virada do ano.
- **Página-container:** `🗂️ Bancos de Orçamento por Ano` (`3d5c45031f738153b0fdf6858d76d740`) — precisa estar **conectada à integração "Agende Aereos App"** (feito 1x em set/2026), pois é nela que os bancos novos nascem (a API do Notion não deixa criar banco direto na raiz do workspace).
- **Duas limitações da própria API do Notion** (não são bug nosso, confirmado até no conector MCP com permissão de usuário completo): não dá pra criar propriedade tipo `status` nem `place` (mapa) via API. Bancos clonados automaticamente usam **`select`** (em vez de `status`) e **`rich_text`** (em vez de `place`) pros campos "Status" e "Endereço" — o server.js já lê/escreve os dois formatos de forma transparente (`propStatusOrcamento`/`lerEnderecoOrcamento`/`propEnderecoOrcamento`). Funciona igual no app, só não tem a mesma carinha visual do banco de 2026.
- **⚠️ Não automatizável — 5 automações do Notion precisam ser recriadas manualmente** toda vez que um ano novo nasce (copiar do banco de 2026, trocando só o banco de destino/origem; a URL do webhook é sempre a mesma). Automações do Notion (botão/regra "quando X muda, chama URL") não são expostas por nenhuma API, nem a interna do conector MCP com permissão de usuário completo — só cria-se na UI. Checklist guiado: skill `/virada-de-ano`. O código dos webhooks é independente de ano (opera pelo pageId recebido; `/webhook-cache-pago` acha o banco de Apresentações do ano pelo `parent` da proposta), então funciona assim que a automação existir:

  | Automação (banco 2026) | Dispara quando | Chama |
  |---|---|---|
  | Propostas → cria Apresentação | Status vira "Aprovado - Aguardando contrato" | `/webhook-proposta-aprovada` |
  | Propostas → avisa elenco do cachê pago | Cachê vira "PAGO" | `/webhook-cache-pago` |
  | Apresentações → sincroniza Google Calendar | Apresentação criada/editada | `/webhook-apresentacao-notion` |
  | Apresentações → avisa elenco/equipe escalada | ELENCO/Produção Liqui/Técnico de Som/Luz preenchidos | `/webhook-apresentacao-escalacao` |
  | Apresentações → avisa saída | Local Saída/Horário de Saída preenchidos | `/webhook-apresentacao-saida` |
- `/orcamento/datas-disponiveis`, `/orcamento/buscar` e `/orcamento/carregar` já buscam em **todos os anos existentes** (não só o atual), via `listarTodosBancosOrcamento()`.

## Lembretes automáticos de WhatsApp pro produtor (set/2026)

Duas mensagens agendadas (fila Notion-backed, nunca `scheduledAt` do Digisac) pro contato em **Produção Liqui** de cada Apresentação, calculadas a partir de `Data da Apresentação` + `Horário Apresentação` (extrai só o horário de início, regex `(\d{1,2})h(\d{2})?` — mesmo padrão do `sincronizarApresentacaoComCalendar`):

- **1h antes:** lembrete de fotos (início/meio/fim), filmar com o celular da Cia. (senha `142536`), contar público.
- **1h depois:** lembrete de preencher o relatório em `apresentacao.ciadoliquidificador.com.br` e subir o vídeo pro YouTube (privado, nome padronizado "Trabalho - Local - Data").

Implementado em `agendarLembretesProdutor(pageId)` (server.js), chamado a partir dos 3 gatilhos que já existem pra Apresentações (`/webhook-apresentacao-notion`, `/webhook-apresentacao-escalacao`, `/webhook-proposta-aprovada`) — `/webhook-apresentacao-escalacao` é o mais confiável, pois dispara exatamente quando Produção Liqui é definida. Idempotente via checkbox `Lembretes Produtor Agendados` (propriedade já adicionada ao molde 2026 e ao teste 2027 — bancos novos herdam automaticamente) **+ trava em memória por pageId** (ver regra 12 acima — sem ela, os 3 gatilhos disparando quase juntos duplicavam a mensagem). Não agenda nada se a apresentação já aconteceu há mais de 2h (edição tardia de registro antigo não deve gerar lembrete fora de hora) ou se o produtor não tem telefone/DDD válido cadastrado no Integrantes.

## Conciliação de recebimentos pelo extrato (portal admin, set/2026)

`pgtAnalisarExtratoRecebimentos` (server.js) cruza os Pix recebidos do extrato Nubank com os Recebimentos `Pendente`. Ordem de tentativa por Pix:

1. Nome bate + um Pendente com valor exato.
2. Nome bate, mas o Pix é a **soma de vários Pendentes** da mesma aluna (duas turmas no mesmo mês, ou dois meses juntos) — `pgtCombinacaoQueSoma`, força bruta em subconjuntos, prefere menos parcelas e meses mais antigos. Cada parte vira um item próprio na tela, com seu valor.
3. Ainda não fechou → amplia pelo **CPF do pagador**: o extrato mostra só o miolo (`•••.688.788-••` = dígitos 4–9), comparado com o `CPF` de Alunas **e** com a coluna `CPF Pagador (Pix)` (texto livre com um ou mais miolos, ex: `217.348 (Paula Mouzinho)`) — `pgtMapaCpfAlunas`. Cobre responsável pagando pela filha num Pix só (Karoline 93 + Maria Flor 207 = 300) e pagador que nem é aluna. Quando a mãe não manda o CPF, basta copiar o miolo do extrato pra `CPF Pagador (Pix)` da criança.
4. Idempotência: `Identificador Pix` (id único do Nubank) gravado ao aplicar + trava "mesma pessoa, mesma data, valor pago ≥ 90% do Pix" pros lançamentos antigos. Subir o mesmo extrato duas vezes (ou um que inclua o mês anterior) não casa de novo — aparece em "🔁 Já aplicados". Isso já aconteceu (set/2026: extrato de agosto casou 16 Set/26) antes da trava existir.

Só o que sobrar cai em "Valor não bate". Se cair, o mais comum é **cadastro errado em Alunas → Valor** (preço do mensal num plano anual/semestral): `gerar-mes` copia esse campo, então corrigir em Alunas e no Pendente do mês. Conferir o histórico de `Valor Pago` dos meses anteriores da aluna antes de mexer.

## Disparos de e-mail (portal admin, set/2026)

Tela `disparos.html` (portal admin) + bloco `PORTAL ADMIN — DISPAROS DE E-MAIL` no server.js. Público vem do **CRM de vendas** no Notion, que fica em outro token: `CRM_NOTION_TOKEN` (mesmo token da triagem de e-mail, integração "Triagem Email CRM") — o `NOTION_TOKEN` normal não enxerga esses bancos. IDs de *database* (não de data source): 📧 Contatos `c7d34260-f3d4-4aac-83c5-cd4ce292c5fa`, 🏛️ Locais `ed6c88ae-05d5-4309-a4b8-76ac23b16327`, 📨 Disparos de E-mail `a3405e65-7376-4cb9-8490-7b2a32f948f1` (log, uma linha por contato por campanha).

- Envio pelo SMTP Locaweb do `contato@cialiquidificador.com.br` (vars `DISPARO_SMTP_USER/PASS/SMTP_HOST/SMTP_PORT`, porta 465). **Um e-mail individual por contato, nunca CCO**, intervalo `DISPARO_INTERVALO_MS` (60s) ±15s. Fila em memória (1 réplica); se o Railway reiniciar no meio, reenviar a campanha **com o mesmo nome** pula quem já consta como Enviado.
- Limites decididos pelo Fábio: **3 MB total** (html + imagens embutidas + anexos, já em base64) e **2 MB por anexo** — estourou, não envia. Imagem colada no editor vira anexo inline `cid:` (Gmail descarta data:URI).
- Rotas exigem sessão admin via header `X-Admin-Token` (`verificarSessao`). Público sempre deduplicado por e-mail (`casasdecultura@` tem 20 linhas) e ignora `Status do E-mail` preenchido (bounce).
- Exclusão pedida pelo Fábio: nunca cadastrar/disparar pra `@semparedescultural.com.br` (Paula Simões, produtora intermediária).

## Apps ativos — mapa completo de FTP (confirmado direto no servidor, set/2026)

**Nunca perguntar "em qual pasta isso vai" — a tabela abaixo já responde.** Pasta local → pasta remota dentro de `/public_html/` na Locaweb. Todas as pastas locais ficam em `~/Public/` (irmãs de `Agende-Aula/`), exceto o Portal Admin.

| Pasta local | Pasta remota (`/public_html/…`) | URL | Tipo |
|---|---|---|---|
| `agende-aereos-app/` | `agende-aereos/` | agende-aereos.ciadoliquidificador.com.br | React (CRA) — agendamento |
| `Agende-Acro-app/` | `agende-acro/` | agende-acro.ciadoliquidificador.com.br | React (CRA) — agendamento |
| `Agende-Infantil-app/` | `agende-infantil/` | agende-infantil.ciadoliquidificador.com.br | React (CRA) — agendamento |
| `Agende-Percussao-app/` | `percussao/` | percussao.ciadoliquidificador.com.br | React (CRA) — matrícula (nome remoto diferente do local!) |
| `agende-commedia-app/` | `commedia/` | commedia.ciadoliquidificador.com.br | React (CRA) — matrícula (nome remoto diferente do local!) |
| `Agende-Ensaio-app/` | `agende-ensaio/` | agende-ensaio.ciadoliquidificador.com.br | HTML estático — Sala de Ensaio |
| `Agende-Yoga-app/` | `agende-yoga/` | agende-yoga.ciadoliquidificador.com.br | HTML estático — booking Yoga |
| `portal-admin-deploy/` | `admin/` | admin.ciadoliquidificador.com.br | HTML estático — Portal Admin |
| `Aluna/` | `aluna/` | aluna.ciadoliquidificador.com.br | HTML estático — Portal Aluna |
| `prof/` | `prof/` | prof.ciadoliquidificador.com.br | HTML estático — Portal Profs |
| `equipe/` | `equipe/` | equipe.ciadoliquidificador.com.br | HTML estático — Portal Artista/Equipe |
| `Presença/` | `presenca/` | presenca.ciadoliquidificador.com.br | HTML estático — chamada de presença |
| `Sub/` | `sub/` | sub.ciadoliquidificador.com.br | HTML estático — substituição de professor |
| `migracao/` | `migracao/` | migracao.ciadoliquidificador.com.br | HTML estático — completar cadastro legado |
| `Matricula/` | `matricula/` | matricula.ciadoliquidificador.com.br | HTML estático — wizard de matrícula |
| `Links/` | `links/` | links.ciadoliquidificador.com.br | HTML estático — link na bio |
| `meditacao/` | `meditacao/` | meditacao.ciadoliquidificador.com.br | HTML estático — curso gratuito |
| `Danças Brasileiras/` | `dancas-brasileiras/` | dancas-brasileiras.ciadoliquidificador.com.br | HTML estático — página de curso |
| `Yoga/` | `yoga/` | yoga.ciadoliquidificador.com.br | HTML estático — página de curso (≠ agende-yoga, que é o booking) |
| `avaliacao-luz/` | `Avaliacao-luz/` | (subpasta, sem subdomínio próprio) | HTML estático — pesquisa de satisfação |
| `Apresentacoes/` | `apresentacao/` | (subpasta) | HTML estático — relatório pós-show (singular no remoto!) |
| `remarcar-residente/` | `remarcar-residente/` | (subpasta) | HTML estático — remarcação Cia Plá |
| `Rersidencia/` (nome local com erro de digitação) | `residencia/` | residencia.ciadoliquidificador.com.br | HTML estático — Residência Artística (nome remoto sem o erro) |
| `Contratos - Professores/` | `contratos-professores/` | contratos-professores.ciadoliquidificador.com.br | HTML estático — assinatura de contrato professor |
| `Aereos/` | `aereos/` | aereos.ciadoliquidificador.com.br | HTML estático — landing/marketing (≠ agende-aereos, que é o booking) |
| `Acro/` | `acro/` | acro.ciadoliquidificador.com.br | HTML estático — landing/marketing |
| `Infantil/` | `infantil/` | infantil.ciadoliquidificador.com.br | HTML estático — landing/marketing |
| `espaço/` | `espaco/` | espaco.ciadoliquidificador.com.br | HTML estático — locação da sala |
| `Reposição/` | `Reposicao/` | (subpasta) | HTML estático — hub de links de reposição |

**Nomes remotos sem acento/cedilha** (Locaweb normaliza): `Reposição`→`Reposicao`, `espaço`→`espaco`, `Presença`→`presenca`, `Danças Brasileiras`→`dancas-brasileiras`. **Nomes remotos que fogem totalmente do nome da pasta local**: Percussão (`Agende-Percussao-app` → `percussao/`) e Commedia (`agende-commedia-app` → `commedia/`) — sem o prefixo "agende-"; Residência (`Rersidencia/` local, com erro de digitação → `residencia/` remoto, sem erro).

**Pendências não confirmadas** (não assumir, perguntar antes de mexer):
- `/public_html/cursos/` está **vazia** no servidor — resquício antigo, sem correspondência local conhecida. Não usar sem confirmar com o Fábio.
- `/public_html/orcamento/` tem um `index.html` de verdade — mas este CLAUDE.md diz que a calculadora de orçamento não fica no FTP (Fábio abre local). Contradição não resolvida: pode ser uma versão antiga esquecida no servidor. **Confirmar com o Fábio antes de sobrescrever ou apagar essa pasta.**

**Deploy de app React (CRA)**: rodar `npm run build` na pasta do app, depois subir o conteúdo de `build/` (não a pasta toda) pro caminho remoto — **sempre limpando o conteúdo antigo de `static/js`/`static/css` antes**, porque o CRA gera nome de arquivo com hash novo a cada build e deploys antigos nunca limpam os hashes velhos (achado real, set/2026: 3 versões antigas acumuladas do bundle principal em produção).

**Importante:** Agende Acro NÃO tem a distinção de reposição/cota que os outros apps (Aéreos/Infantil/Yoga) têm. Não replicar lógica de cota lá sem pedido explícito.

**Portal Admin:** fonte local em `~/Public/portal-admin-deploy/` (HTML estático, telas carregadas via iframe dentro de `admin.html`). Instalável como PWA no celular (manifest.json + sw.js, set/2026) — no iPhone precisa ser pelo Safari (Compartilhar → Adicionar à Tela de Início); Chrome iOS não expõe essa opção.

## Portais (Profs + Aluna) — jul/2026

Login por CPF + OTP único via WhatsApp. Sessão ativa 10min sem pedir novo código (armazenada em memória no server.js: `sessaoStore`/`otpStore`, chaves `portalAlunaSessao`/`portalProfessorSessao`).

- **Portal Profs**: Minhas Turmas, Feriados, Rendimento do Mês, Meu Contrato (docência), Meus Dados, Mural de Avisos.
- **Portal Aluna**: Meu Contrato (por matrícula), Minha Presença, Remarcar Aula (deep-link `?reposicao=1` pros apps externos — Acro ainda não tem isso), Mudar de Turma, Mudar de Plano, Contato e Saúde, Meus Dados, Mural de Avisos.
- **Conhecido:** dados nesses portais podem ficar "presos" em cache de sessão por até 10min — se algo não aparecer logo após uma mudança, testar logout/login antes de assumir bug real. (Decisão tomada: não vale a pena mexer nisso, é edge case de teste, não de uso real.)

## Cota de reposição (em desenvolvimento, jul/2026)

**Não é um contador de mês-calendário.** É um sistema de créditos individuais com **janela rolante de 30 dias**:

- Toda Falta registrada em `Presenças 2026` gera um crédito em `🔄 Reposições Solicitadas` com `Status` = "Aberto". `Prazo Limite` depende do Plano (set/2026):
  - **Mensal:** Data da falta + 30 dias (não tem ciclo fixo, renova mês a mês).
  - **Semestral/Anual:** TODOS os créditos do mesmo ciclo (semestre/ano) vencem juntos, 30 dias após o término desse ciclo — não 30 dias após cada falta individual. Ciclo é recorrente a partir de `Data/Hora Aceite Contrato` (não usa `Vencimento do Contrato` direto, que é fixo na 1ª fidelidade e não avança nas renovações automáticas — `calcularProximoTerminoCiclo`/`calcularPrazoLimiteCredito`). Se `Data/Hora Aceite Contrato` estiver vazio (gap de dado legado), cai no fallback igual ao Mensal.
  - `verificarCotaReposicao` extrai a data real da falta do `Título` do crédito (regex `Falta (\d{4}-\d{2}-\d{2})`), nunca de `Prazo Limite - 30` — essa conta só valia quando prazo=falta+30 sempre, o que não é mais verdade pra Semestral/Anual.
- Cota = quantos créditos "Aberto" a aluna pode ter ao mesmo tempo, por modalidade: 1x/semana → 1, 2x/semana → 2, "Acordo" → 1 (fallback, é caso raro de acordo de pagamento, não de frequência real).
- Ao usar um crédito (agendar a reposição), Status vira "Usado" — isso libera vaga pra próxima falta em espera, mesmo que ainda dentro do mês.
- Créditos com `Prazo Limite` vencido e ainda "Aberto" devem ser tratados como "Expirado" (lazy expiration, checar na hora da consulta, sem precisar de cron).
- Ver prompt completo de implementação em `prompt_cota_reposicao_v2.md` (histórico de decisão no chat "Cia do Liquidificador" do Claude.ai).
- **A cota (`calcularCotaReposicao`) diferencia só por Frequência (1x/2x semana), nunca por Plano** (Mensal/Semestral/Anual) — conferido contra a Cláusula Sétima do contrato de matrícula (`montarTextoContratoMatricula`), que também só diferencia por frequência. Se um dia pedirem diferenciação por plano, é regra nova, não um gap a "descobrir" no contrato.
- **`podeAgendar` NUNCA deve bloquear por excesso de créditos acumulados** (`creditos.length > cota`) — só bloqueia quando não há nenhum crédito válido (`creditos.length === 0`). Já existiu um bug (set/2026, aluna Maíra Bombachini) onde `creditos.length > cota` bloqueava o agendamento, mas agendar é a ÚNICA forma de reduzir créditos abertos — virava um beco sem saída permanente assim que a aluna acumulava mais faltas que a cota (ex: 2 faltas seguidas sem repor entre elas). Sempre consome o crédito mais antigo primeiro (`creditos.sort` por Prazo Limite); a janela de 30 dias já limita o acúmulo por conta própria.

## Fluxo de trabalho

- Fábio não testa localmente fora do que é explicitamente pedido. Todo deploy via Railway.
- FTP sempre em `/public_html/[subdomain]/` — nunca a raiz do FTP.
- Terminal output é colado direto no chat do Claude.ai pra interpretação; prefere confirmação concisa a explicação longa.
- `sed` e scripts Python3 são preferidos a `nano` pra edições em lote (nano já corrompeu arquivo). Scripts de uso único (`patch_*.py`, migrações, correções pontuais) ficam em `arquivo/` (ignorado pelo git), nunca soltos na raiz.
- Planejamento, operações no Notion (criar bancos/colunas) e prompts pro Claude Code continuam no Claude.ai (chat "Cia do Liquidificador"). Edição direta e grande no server.js é feita aqui no Claude Code.
