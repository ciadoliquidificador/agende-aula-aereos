# Cia do Liquidificador — Contexto do Projeto

> Este arquivo é lido automaticamente no início de toda sessão do Claude Code.
> Mantenha atualizado conforme o projeto evolui — é a memória de longo prazo do projeto.

## Sobre o negócio

Cia do Liquidificador é uma escola de artes cênicas operando como **Liquidificador Produções Artísticas / Cristiane Socci Leonel - ME** (CNPJ: 28.398.119/0001-89), em Rua Dr. Carvalho de Mendonça, 67 — Campos Elíseos, São Paulo/SP. **Fábio Spila** administra toda a infraestrutura técnica.

**Ano atual: 2026.** Sempre usar 2026 em datas, contratos, nomes de bancos e qualquer documento gerado.

## Stack

- **Frontend:** React apps hospedados na Locaweb via FTP, tema visual creme+vinho (Playfair Display + Inter)
- **Backend:** Um único proxy Node.js/Express compartilhado no Railway (`~/Public/Agende-Aula/server.js`, 4000+ linhas)
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

## Notion — regras e IDs importantes

- **PAT tokens (`ntn_`) são escopados por workspace**, não dão acesso automático a tudo. Bancos herdados via página-pai ("Added via...") são pouco confiáveis via API — só "Selected manually" funciona de forma consistente. Às vezes é preciso um token novo por banco.
- **`ALTER COLUMN ... RENAME`** não é suportado — workaround é `ADD COLUMN "Nome Novo" TYPE` e deixar a coluna antiga sem uso.
- **`ALTER COLUMN SET SELECT(...)`** precisa incluir TODAS as opções já existentes, senão elas somem.
- **Notion Forms:** perguntas e descrições de cabeçalho não podem ser setadas via API — sempre manual na UI do Notion.
- **Money fields:** sempre `number_format: real` (nunca dollar).
- Bancos principais:
  - `Alunas` — `aee12f7f-8cb9-4ee2-80ba-1bcb06d9eda0` (data source `41bb69c4-2d18-4c81-9e43-e60c5f4033f6`)
  - `Presenças 2026` — data source `8365a940-b386-401b-bedb-d26dfff2415e`
  - `🔄 Reposições Solicitadas` — `dde8519e6e0f4157b2bb56b545e2ef84` (data source `53acaa29-45c5-4ace-99b5-8d1149a96e9d`)
  - `📄 Contratos — Professores` — `f1b934a8100143019ac7f5f877c405f1`
  - `📎 Aditivos — Contratos Professores` — `9b35319a73854e318cd9efc6497bfb0e`
  - `📢 Mural de Avisos` — `c45786e213ff463f8558054b2f787a69`
  - `👥 Professores — Cadastro` — `728021ad4c58466db1dd5ab112ada252`

## Apps ativos

| App | URL | FTP path |
|-----|-----|----------|
| Agende Aéreos | agende-aereos.ciadoliquidificador.com.br | /public_html/agende-aereos/ |
| Agende Acro | agende-acro.ciadoliquidificador.com.br | /public_html/agende-acro/ |
| Agende Infantil | agende-infantil.ciadoliquidificador.com.br | /public_html/agende-infantil/ |
| Percussão | percussao.ciadoliquidificador.com.br | /public_html/percussao/ |
| Sala de Ensaio | agende-ensaio.ciadoliquidificador.com.br | (integrado no server.js) |
| Links/Cursos | links.ciadoliquidificador.com.br | (HTML estático) |
| Presença | presenca.ciadoliquidificador.com.br | — |
| Portal Profs | prof.ciadoliquidificador.com.br | — |
| Portal Aluna | aluna.ciadoliquidificador.com.br | — |

**Importante:** Agende Acro NÃO tem a distinção de reposição/cota que os outros apps (Aéreos/Infantil/Yoga) têm. Não replicar lógica de cota lá sem pedido explícito.

## Portais (Profs + Aluna) — jul/2026

Login por CPF + OTP único via WhatsApp. Sessão ativa 10min sem pedir novo código (armazenada em memória no server.js: `sessaoStore`/`otpStore`, chaves `portalAlunaSessao`/`portalProfessorSessao`).

- **Portal Profs**: Minhas Turmas, Feriados, Rendimento do Mês, Meu Contrato (docência), Meus Dados, Mural de Avisos.
- **Portal Aluna**: Meu Contrato (por matrícula), Minha Presença, Remarcar Aula (deep-link `?reposicao=1` pros apps externos — Acro ainda não tem isso), Mudar de Turma, Mudar de Plano, Contato e Saúde, Meus Dados, Mural de Avisos.
- **Conhecido:** dados nesses portais podem ficar "presos" em cache de sessão por até 10min — se algo não aparecer logo após uma mudança, testar logout/login antes de assumir bug real. (Decisão tomada: não vale a pena mexer nisso, é edge case de teste, não de uso real.)

## Cota de reposição (em desenvolvimento, jul/2026)

**Não é um contador de mês-calendário.** É um sistema de créditos individuais com **janela rolante de 30 dias**:

- Toda Falta registrada em `Presenças 2026` gera um crédito em `🔄 Reposições Solicitadas` com `Prazo Limite` = Data da falta + 30 dias, `Status` = "Aberto".
- Cota = quantos créditos "Aberto" a aluna pode ter ao mesmo tempo, por modalidade: 1x/semana → 1, 2x/semana → 2, "Acordo" → 1 (fallback, é caso raro de acordo de pagamento, não de frequência real).
- Ao usar um crédito (agendar a reposição), Status vira "Usado" — isso libera vaga pra próxima falta em espera, mesmo que ainda dentro do mês.
- Créditos com `Prazo Limite` vencido e ainda "Aberto" devem ser tratados como "Expirado" (lazy expiration, checar na hora da consulta, sem precisar de cron).
- Ver prompt completo de implementação em `prompt_cota_reposicao_v2.md` (histórico de decisão no chat "Cia do Liquidificador" do Claude.ai).

## Fluxo de trabalho

- Fábio não testa localmente fora do que é explicitamente pedido. Todo deploy via Railway.
- FTP sempre em `/public_html/[subdomain]/` — nunca a raiz do FTP.
- Terminal output é colado direto no chat do Claude.ai pra interpretação; prefere confirmação concisa a explicação longa.
- `sed` e scripts Python3 são preferidos a `nano` pra edições em lote (nano já corrompeu arquivo).
- Planejamento, operações no Notion (criar bancos/colunas) e prompts pro Claude Code continuam no Claude.ai (chat "Cia do Liquidificador"). Edição direta e grande no server.js é feita aqui no Claude Code.
