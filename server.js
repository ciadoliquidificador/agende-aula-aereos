const express = require('express');
const fetch = require('node-fetch');
const app = express();

app.use(express.json());

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

const DIGISAC_URL = 'https://ciadoliquidificador.digisac.biz/api/v1/messages';
const DIGISAC_TOKEN = 'c2e0a3ac3ae20585924b2dbc133d68d779770199';
const SERVICE_ID = '012587f9-21ea-4143-9005-c0fbdf109f05';
const USER_ID = 'b0bb99db-a668-403a-af70-efc1d4a7259a';
const NOTION_TOKEN = process.env.NOTION_TOKEN;
const ALUNAS_DB = process.env.NOTION_DATABASE_ID || 'aee12f7f-8cb9-4ee2-80ba-1bcb06d9eda0';

async function enviarWhatsApp(numero, texto) {
  const response = await fetch(DIGISAC_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${DIGISAC_TOKEN}` },
    body: JSON.stringify({ text: texto, type: 'chat', serviceId: SERVICE_ID, number: numero, userId: USER_ID, origin: 'bot' }),
  });
  if (!response.ok) { const t = await response.text(); throw new Error(`Digisac ${response.status}: ${t}`); }
  return response.json();
}

app.get('/health', (req, res) => res.json({ ok: true }));

// Retorna turmas com alunas ativas + experimentais (para calcular vagas)
app.get('/turmas', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        filter: {
          and: [
            {
              or: [
                { property: 'Status', select: { equals: 'Ativa' } },
                { property: 'Status', select: { equals: 'Experimental' } },
              ]
            },
            { property: 'Modalidade', select: { equals: 'Aéreos' } },
          ]
        },
        page_size: 100,
      }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    const data = await response.json();

    const map = {};
    for (const page of data.results) {
      const props = page.properties;
      const nome = props.Nome?.title?.[0]?.plain_text || '';
      const turma = props.Turma?.select?.name || props.Turma?.rich_text?.[0]?.plain_text || '';
      const professor = props.Professor?.select?.name || props.Professor?.rich_text?.[0]?.plain_text || '';
      const status = props.Status?.select?.name || '';
      if (!turma) continue;
      if (!map[turma]) {
        map[turma] = { id: `t_${turma.replace(/[\s/:]/g, '_')}`, nome: `${turma} - Prof. ${professor}`, professor, alunas: [], experimentais: [] };
      }
      if (status === 'Experimental') {
        if (nome && !map[turma].experimentais.includes(nome)) map[turma].experimentais.push(nome);
      } else {
        if (nome && !map[turma].alunas.includes(nome)) map[turma].alunas.push(nome);
      }
    }
    return res.json({ ok: true, turmas: Object.values(map) });
  } catch (err) {
    return res.json({ ok: false, erro: err.message, turmas: [] });
  }
});

// Cria inscrição experimental no Notion
app.post('/inscricao', async (req, res) => {
  try {
    const { nome, telefone, turma, professor, dia, horario, data } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });

    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        parent: { database_id: ALUNAS_DB },
        properties: {
          Nome: { title: [{ text: { content: nome } }] },
          Contato: { phone_number: telefone },
          Turma: { select: { name: turma } },
          Professor: { select: { name: professor || '' } },
          Dia: { select: { name: dia || '' } },
          Horário: { select: { name: horario || '' } },
          Modalidade: { select: { name: 'Aéreos' } },
          Status: { select: { name: 'Experimental' } },
          Observações: { rich_text: [{ text: { content: `Aula experimental agendada para ${data || ''}` } }] },
        },
      }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    return res.json({ ok: true });
  } catch (err) {
    return res.json({ ok: false, erro: err.message });
  }
});

app.post('/', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/agendar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/notificar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/lembrete', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Proxy rodando na porta ${PORT}`));
