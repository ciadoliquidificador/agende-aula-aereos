const express = require('express');
const fetch = require('node-fetch');
const app = express();

app.use(express.json({ limit: '60mb' }));

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

const DIGISAC_BASE = 'https://ciadoliquidificador.digisac.biz/api/v1';
const DIGISAC_TOKEN = process.env.DIGISAC_TOKEN;
const SERVICE_ID = '012587f9-21ea-4143-9005-c0fbdf109f05';
const USER_ID = 'b0bb99db-a668-403a-af70-efc1d4a7259a';
const NOTION_TOKEN = process.env.NOTION_TOKEN;

// ============================================================
// ONEDRIVE (Microsoft Graph) - Upload de fotos das apresentacoes
// ============================================================
const MS_TENANT_ID = process.env.MICROSOFT_TENANT_ID;
const MS_CLIENT_ID = process.env.MICROSOFT_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET;
const MS_USER = 'fspila@ciadoliquidificador.onmicrosoft.com';
const MS_FOLDER_ID = '01WKZ3MCJVE2PAHHDUPZHYGSWRT2QQEQ2A';

async function getMicrosoftToken() {
  const params = new URLSearchParams();
  params.append('client_id', MS_CLIENT_ID);
  params.append('client_secret', MS_CLIENT_SECRET);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('grant_type', 'client_credentials');

  const r = await fetch('https://login.microsoftonline.com/' + MS_TENANT_ID + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });
  if (!r.ok) { const t = await r.text(); throw new Error('Auth Microsoft: ' + t); }
  const data = await r.json();
  return data.access_token;
}

async function criarOuObterSubpasta(token, nomePasta) {
  const nomeSeguro = nomePasta.replace(/[<>:"/\\|?*]/g, '-').slice(0, 100);

  const createUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + MS_FOLDER_ID + '/children';
  const createResp = await fetch(createUrl, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: nomeSeguro, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
  });

  if (createResp.ok) {
    const data = await createResp.json();
    return data.id;
  }

  if (createResp.status === 409) {
    const getUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + MS_FOLDER_ID + ':/' + encodeURIComponent(nomeSeguro);
    const getResp = await fetch(getUrl, { headers: { 'Authorization': 'Bearer ' + token } });
    if (getResp.ok) {
      const data = await getResp.json();
      return data.id;
    }
  }

  const t = await createResp.text();
  throw new Error('Criar subpasta: ' + t);
}

async function uploadFotoParaPasta(token, folderId, base64Data, filename) {
  const base64Clean = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const buffer = Buffer.from(base64Clean, 'base64');

  const uploadUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + folderId + ':/' + encodeURIComponent(filename) + ':/content';

  const r = await fetch(uploadUrl, {
    method: 'PUT',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/octet-stream' },
    body: buffer,
  });
  if (!r.ok) { const t = await r.text(); throw new Error('Upload foto: ' + t); }
  return await r.json();
}

async function criarLinkCompartilhamento(token, itemId) {
  const shareResp = await fetch('https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + itemId + '/createLink', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ type: 'view', scope: 'anonymous' }),
  });
  if (!shareResp.ok) { const t = await shareResp.text(); throw new Error('Criar link: ' + t); }
  const shareData = await shareResp.json();
  return shareData.link.webUrl;
}

async function uploadFotoOneDrive(token, base64Data, filename) {
  const base64Clean = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const buffer = Buffer.from(base64Clean, 'base64');

  const uploadUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + MS_FOLDER_ID + ':/' + encodeURIComponent(filename) + ':/content';

  const r = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/octet-stream',
    },
    body: buffer,
  });
  if (!r.ok) { const t = await r.text(); throw new Error('OneDrive upload: ' + t); }
  const data = await r.json();

  const shareResp = await fetch('https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + data.id + '/createLink', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ type: 'view', scope: 'anonymous' }),
  });
  if (shareResp.ok) {
    const shareData = await shareResp.json();
    return shareData.link.webUrl;
  }
  return data.webUrl;
}
const ALUNAS_DB = process.env.NOTION_DATABASE_ID || 'aee12f7f-8cb9-4ee2-80ba-1bcb06d9eda0';

const digisacHeaders = {
  'Content-Type': 'application/json',
  'Authorization': `Bearer ${DIGISAC_TOKEN}`
};

async function getOrCreateContactId(numero) {
  const num = numero.replace(/\D/g, '');
  const numBr = num.startsWith('55') ? num : `55${num}`;
  const searchRes = await fetch(`${DIGISAC_BASE}/contacts?number=${numBr}&serviceId=${SERVICE_ID}`, { headers: digisacHeaders });
  if (searchRes.ok) {
    const searchData = await searchRes.json();
    if (searchData.data && searchData.data.length > 0) {
      const found = searchData.data.find(c => c.data?.number === numBr || c.data?.number === num);
      if (found) return found.id;
      // não usa fallback — segue para criar contato
    }
  }
  const createRes = await fetch(`${DIGISAC_BASE}/contacts`, {
    method: 'POST',
    headers: digisacHeaders,
    body: JSON.stringify({ serviceId: SERVICE_ID, number: numBr, name: numBr })
  });
  if (!createRes.ok) { const t = await createRes.text(); throw new Error(`Digisac criar contato ${createRes.status}: ${t}`); }
  const created = await createRes.json();
  return created.id || created.data?.id;
}

async function enviarWhatsApp(numero, texto) {
  const contactId = await getOrCreateContactId(numero);
  const response = await fetch(`${DIGISAC_BASE}/messages`, {
    method: 'POST',
    headers: digisacHeaders,
    body: JSON.stringify({ text: texto, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot' })
  });
  if (!response.ok) { const t = await response.text(); throw new Error(`Digisac ${response.status}: ${t}`); }
  return response.json();
}

function calcularProximoHorarioComercial() {
  const agora = new Date();
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Sao_Paulo',
    hour: 'numeric', hour12: false,
    weekday: 'short',
  });
  const diasMap = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };

  const parts = fmt.formatToParts(agora);
  const hora = parseInt(parts.find(p => p.type === 'hour').value, 10);
  const diaSemana = diasMap[parts.find(p => p.type === 'weekday').value];

  const dentroHorario = hora >= 8 && hora < 18 && diaSemana >= 1 && diaSemana <= 5;
  if (dentroHorario) return null;

  // Avanca hora a hora ate encontrar 8h de um dia util (Brasilia)
  let candidato = new Date(agora);
  for (let i = 0; i < 24 * 8; i++) {
    candidato = new Date(candidato.getTime() + 60 * 60000);
    const partsC = fmt.formatToParts(candidato);
    const horaC = parseInt(partsC.find(p => p.type === 'hour').value, 10);
    const diaC = diasMap[partsC.find(p => p.type === 'weekday').value];
    if (horaC === 8 && diaC >= 1 && diaC <= 5) {
      return candidato;
    }
  }
  return null;
}

async function enviarWhatsAppComHorarioComercial(numero, texto) {
  const agendamento = calcularProximoHorarioComercial();
  console.log('[horario-comercial] agora=' + new Date().toISOString() + ' agendamento=' + (agendamento ? agendamento.toISOString() : 'null (envio imediato)'));
  const contactId = await getOrCreateContactId(numero);
  if (!contactId) throw new Error('Contato nao encontrado');
  const body = { text: texto, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot' };
  if (agendamento) body.scheduledAt = agendamento.toISOString();
  const response = await fetch(DIGISAC_BASE + '/messages', {
    method: 'POST', headers: digisacHeaders, body: JSON.stringify(body),
  });
  if (!response.ok) { const t = await response.text(); throw new Error('Digisac ' + response.status + ': ' + t); }
  return { agendado: !!agendamento };
}

app.get('/health', (req, res) => res.json({ ok: true }));

app.get('/turmas', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { and: [{ or: [{ property: 'Status', select: { equals: 'Ativa' } }, { property: 'Status', select: { equals: 'Experimental' } }] }, { property: 'Modalidade', select: { equals: 'Aéreos' } }] }, page_size: 100 }),
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
      if (!map[turma]) map[turma] = { id: `t_${turma.replace(/[\s:]/g, '_')}`, nome: `${turma} - Prof. ${professor}`, professor, alunas: [], experimentais: [] };
      if (status === 'Experimental') { if (nome && !map[turma].experimentais.includes(nome)) map[turma].experimentais.push(nome); }
      else { if (nome && !map[turma].alunas.includes(nome)) map[turma].alunas.push(nome); }
    }
    return res.json({ ok: true, turmas: Object.values(map) });
  } catch (err) { return res.json({ ok: false, erro: err.message, turmas: [] }); }
});

app.post('/inscricao', async (req, res) => {
  try {
    const { nome, telefone, turma, professor, dia, horario, data } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });
    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: { Nome: { title: [{ text: { content: nome } }] }, Contato: { phone_number: telefone }, Turma: { select: { name: turma } }, Professor: { select: { name: professor || '' } }, Dia: { select: { name: dia || '' } }, Horário: { select: { name: horario || '' } }, Modalidade: { select: { name: 'Aéreos' } }, Status: { select: { name: 'Experimental' } }, Observações: { rich_text: [{ text: { content: `Aula experimental agendada para ${data || ''}` } }] } } }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    return res.json({ ok: true });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

app.post('/enviar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/agendar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/notificar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/lembrete', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Proxy rodando na porta ${PORT}`));

const nomeParaId = {
  // Formato novo (padrão)
  "Segunda 18h": "t1",
  "Segunda 19h": "t2",
  "Sexta 18h":   "t3",
  "Terça 8h":    "t4",
  "Terça 9h":    "t5",
  "Quarta 18h":  "t6",
  "Quarta 19h":  "t7",
  "Quinta 8h":   "t8",
  // Formato legado (registros antigos)
  "Segunda 18h – Prof. Gabi":      "t1",
  "Segunda 19h – Prof. Gabi":      "t2",
  "Sexta 18h – Prof. Gabi":        "t3",
  "Terça 8h – Prof. Talita":       "t4",
  "Terça 9h – Prof. Talita":       "t5",
  "Quarta 18h – Prof. Gustra":     "t6",
  "Quarta 19h – Prof. Gustra":     "t7",
  "Quinta 8h – Prof. Guilherme":   "t8",
};

app.get('/vagas-ocupadas', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Status', select: { equals: 'Experimental' } }, page_size: 100 }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    const data = await response.json();
    const ocupadas = data.results.map(p => {
      const props = p.properties;
      const turma = props.Turma?.select?.name || '';
      const obs = props.Observações?.rich_text?.[0]?.plain_text || '';
      const dataMatch = obs.match(/(\d{4}-\d{2}-\d{2})/);
      const data = dataMatch ? dataMatch[1] : '';
      const turmaId = nomeParaId[turma];
      return { turmaId, data };
    }).filter(v => v.turmaId && v.data);
    return res.json({ ok: true, ocupadas });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

// ── Fila de mensagens agendadas ───────────────────────────────────────────────
const filaMensagens = [];

app.post('/agendar-mensagem', async (req, res) => {
  const { numero, texto, enviarEm } = req.body;
  if (!numero || !texto || !enviarEm) return res.json({ ok: false, erro: 'Campos obrigatorios.' });
  try {
    const contactId = await getOrCreateContactId(numero);
    if (!contactId) return res.json({ ok: false, erro: 'Contato nao encontrado.' });
    const response = await fetch(DIGISAC_BASE + '/messages', {
      method: 'POST', headers: digisacHeaders,
      body: JSON.stringify({
        text: texto, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot',
        scheduledAt: new Date(enviarEm).toISOString(),
      }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error('Digisac ' + response.status + ': ' + t); }
    return res.json({ ok: true });
  } catch (err) {
    console.error('[agendar-mensagem] erro:', err.message);
    return res.json({ ok: false, erro: err.message });
  }
});

// ── Circo Acrobacia ───────────────────────────────────────────────────────────
const ACRO_TURMA_NOME = 'Segunda 10h – Circo Acrobacia';

app.post('/inscricao-acro', async (req, res) => {
  try {
    const { nome, telefone, turma, professor, dia, horario, data } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });
    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: { Nome: { title: [{ text: { content: nome } }] }, Contato: { phone_number: telefone }, Turma: { select: { name: turma } }, Professor: { select: { name: 'André' } }, Dia: { select: { name: dia || 'Segunda' } }, Horário: { select: { name: horario || '10:00' } }, Modalidade: { select: { name: 'Circo - Acrobacia' } }, Status: { select: { name: 'Experimental' } }, Observações: { rich_text: [{ text: { content: `Aula experimental agendada para ${data || ''}` } }] } } }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }

    // Lembrete 24h para a aluna (confirmacao ja e feita pelo frontend)
    try {
      if (data) {
        const numLimpo = (telefone || '').replace(/\D/g, '');
        const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
        const primeiroNome = nome.split(' ')[0];
        const horarioFmt = (horario || '10:00').replace(':00', 'h');

        const dataAula = new Date(data + 'T' + (horario || '10:00') + ':00-03:00');
        const lembrete = new Date(dataAula);
        lembrete.setDate(lembrete.getDate() - 1);
        lembrete.setHours(8, 0, 0, 0);
        const msgLembrete = `Olá, ${primeiroNome}! 🤸\n\nLembrando que amanhã você tem aula experimental de Acrobacias!\n\n⏰ Horário: ${horarioFmt}\n📍 Rua Dr. Carvalho de Mendonça, 67 — Campos Elíseos`;
        const contactId = await getOrCreateContactId(numBr);
        if (contactId) {
          await fetch(DIGISAC_BASE + '/messages', {
            method: 'POST', headers: digisacHeaders,
            body: JSON.stringify({ text: msgLembrete, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot', scheduledAt: lembrete.toISOString() }),
          });
        }
      }
    } catch(e) { console.error('[acro] erro ao agendar lembrete:', e.message); }

    return res.json({ ok: true });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

app.get('/vagas-ocupadas-acro', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { and: [{ property: 'Status', select: { equals: 'Experimental' } }, { property: 'Modalidade', select: { equals: 'Circo - Acrobacia' } }] }, page_size: 100 }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    const data = await response.json();
    const ocupadas = data.results.map(p => {
      const props = p.properties;
      const obs = props.Observações?.rich_text?.[0]?.plain_text || '';
      const dataMatch = obs.match(/(\d{4}-\d{2}-\d{2})/);
      const data = dataMatch ? dataMatch[1] : '';
      return { turmaId: 'a1', data };
    }).filter(v => v.data);
    return res.json({ ok: true, ocupadas });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

// ── Circo Infantil ────────────────────────────────────────────────────────────
app.post('/inscricao-infantil', async (req, res) => {
  try {
    const { nome, telefone, turma, professor, dia, horario, data, tipo } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });
    const status = tipo === 'reposicao' ? 'Ativo' : 'Experimental';
    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: { Nome: { title: [{ text: { content: nome } }] }, Contato: { phone_number: telefone }, Turma: { select: { name: turma } }, Professor: { select: { name: 'Titzi' } }, Dia: { select: { name: dia || 'Terça' } }, Horário: { select: { name: horario || '18:00' } }, Modalidade: { select: { name: 'Circo Infantil' } }, Status: { select: { name: status } }, Observações: { rich_text: [{ text: { content: `Agendado para ${data || ''}` } }] } } }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    return res.json({ ok: true });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

app.get('/vagas-ocupadas-infantil', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { and: [{ property: 'Status', select: { equals: 'Experimental' } }, { property: 'Modalidade', select: { equals: 'Circo Infantil' } }] }, page_size: 100 }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    const data = await response.json();
    const ocupadas = data.results.map(p => {
      const obs = p.properties.Observações?.rich_text?.[0]?.plain_text || '';
      const turma = p.properties.Turma?.select?.name || '';
      const dataMatch = obs.match(/(\d{4}-\d{2}-\d{2})/);
      const dataAula = dataMatch ? dataMatch[1] : '';
      const turmaId = turma.includes('Terça') ? 'ci1' : turma.includes('Quarta') ? 'ci2' : '';
      return { turmaId, data: dataAula };
    }).filter(v => v.turmaId && v.data);
    return res.json({ ok: true, ocupadas });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

// ── Circo Infantil ────────────────────────────────────────────────────────────
app.post('/inscricao-infantil', async (req, res) => {
  try {
    const { nome, telefone, turma, professor, dia, horario, data, tipo } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });
    const status = tipo === 'reposicao' ? 'Ativo' : 'Experimental';
    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: { Nome: { title: [{ text: { content: nome } }] }, Contato: { phone_number: telefone }, Turma: { select: { name: turma } }, Professor: { select: { name: 'Titzi' } }, Dia: { select: { name: dia || 'Terça' } }, Horário: { select: { name: horario || '18:00' } }, Modalidade: { select: { name: 'Circo Infantil' } }, Status: { select: { name: status } }, Observações: { rich_text: [{ text: { content: `Agendado para ${data || ''}` } }] } } }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    return res.json({ ok: true });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

app.get('/vagas-ocupadas-infantil', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { and: [{ property: 'Status', select: { equals: 'Experimental' } }, { property: 'Modalidade', select: { equals: 'Circo Infantil' } }] }, page_size: 100 }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    const data = await response.json();
    const ocupadas = data.results.map(p => {
      const obs = p.properties.Observações?.rich_text?.[0]?.plain_text || '';
      const turma = p.properties.Turma?.select?.name || '';
      const dataMatch = obs.match(/(\d{4}-\d{2}-\d{2})/);
      const dataAula = dataMatch ? dataMatch[1] : '';
      const turmaId = turma.includes('Terça') ? 'ci1' : turma.includes('Quarta') ? 'ci2' : '';
      return { turmaId, data: dataAula };
    }).filter(v => v.turmaId && v.data);
    return res.json({ ok: true, ocupadas });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

// ── Percussão Coletiva ────────────────────────────────────────────────────────
const PERCUSSAO_DB = '94b8c7ff-792c-4360-990d-f1d892cb78a3';
const PERCUSSAO_TOKEN = process.env.PERCUSSAO_TOKEN;

app.get('/vagas-percussao', async (req, res) => {
  try {
    const response = await fetch(`https://api.notion.com/v1/databases/${PERCUSSAO_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${PERCUSSAO_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Status', select: { does_not_equal: 'Cancelado' } }, page_size: 100 }),
    });
    const data = await response.json();
    const inscritos = data.results ? data.results.length : 0;
    return res.json({ ok: true, inscritos, disponiveis: Math.max(0, 20 - inscritos) });
  } catch (err) { return res.json({ ok: false, erro: err.message, disponiveis: 20 }); }
});

app.post('/inscricao-percussao', async (req, res) => {
  try {
    const { nomeCompleto, nomeSocial, rg, cpf, telefone, email, rua, numero, bairro, cep, cidade, estado, tocaPercussao, quaisInstrumentos, temInstrumentos, outrosInstrumentosToca, outrosInstrumentosTem, formaPagamento, assinatura, observacoes } = req.body;
    if (!nomeCompleto || !cpf || !telefone) return res.json({ ok: false, erro: 'Campos obrigatórios ausentes.' });
    const response = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${PERCUSSAO_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: PERCUSSAO_DB }, properties: {
        'Nome Completo': { title: [{ text: { content: nomeCompleto } }] },
        'Nome Social': { rich_text: [{ text: { content: nomeSocial || '' } }] },
        'RG': { rich_text: [{ text: { content: rg || '' } }] },
        'CPF': { rich_text: [{ text: { content: cpf } }] },
        'Telefone': { phone_number: telefone },
        'E-mail': { email: email },
        'Rua': { rich_text: [{ text: { content: rua || '' } }] },
        'Número': { rich_text: [{ text: { content: numero || '' } }] },
        'Bairro': { rich_text: [{ text: { content: bairro || '' } }] },
        'CEP': { rich_text: [{ text: { content: cep || '' } }] },
        'Cidade': { rich_text: [{ text: { content: cidade || '' } }] },
        'Estado': { rich_text: [{ text: { content: estado || '' } }] },
        'Toca percussão?': { select: { name: tocaPercussao || '' } },
        'Quais instrumentos toca': { multi_select: (quaisInstrumentos || []).map(n => ({ name: n })) },
        'Tem instrumentos': { multi_select: (temInstrumentos || []).map(n => ({ name: n })) },
        'Forma de Pagamento': { select: { name: formaPagamento } },
        'Aceite dos Termos': { select: { name: 'Sim - li e aceito os termos acima' } },
        'Assinatura': { rich_text: [{ text: { content: assinatura || '' } }] },
        'Observações': { rich_text: [{ text: { content: observacoes || '' } }] },
        'Outros instrumentos que toca': { rich_text: [{ text: { content: outrosInstrumentosToca || '' } }] },
        'Outros instrumentos que tem': { rich_text: [{ text: { content: outrosInstrumentosTem || '' } }] },
        'Status': { select: { name: 'Inscrito' } },
      }}),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }
    return res.json({ ok: true });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

// ============================================================
// COMMEDIA DELL'ARTE
// ============================================================
const NOTION_TOKEN_COMMEDIA = process.env.NOTION_TOKEN_COMMEDIA;
const NOTION_DB_COMMEDIA = '1cdeb67a3652412cacc38347a2c6c5ba';
const MAX_VAGAS_COMMEDIA = 18;

app.get('/vagas-commedia', async (req, res) => {
  try {
    const r = await fetch(`https://api.notion.com/v1/databases/${NOTION_DB_COMMEDIA}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN_COMMEDIA}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Status', select: { does_not_equal: 'Cancelado' } } }),
    });
    const data = await r.json();
    const ocupadas = data.results ? data.results.length : 0;
    res.json({ ok: true, disponiveis: Math.max(0, MAX_VAGAS_COMMEDIA - ocupadas) });
  } catch (err) {
    res.json({ ok: false, disponiveis: MAX_VAGAS_COMMEDIA });
  }
});

app.post('/inscricao-commedia', async (req, res) => {
  const { nomeCompleto, nomeSocial, rg, cpf, telefone, email, rua, numero, complemento, bairro, cep, cidade, estado, formaPagamento, assinatura, observacoes } = req.body;
  if (!nomeCompleto || !cpf || !telefone || !email || !formaPagamento || !assinatura) return res.status(400).json({ error: 'Campos obrigatórios faltando' });
  const numLimpo = telefone.replace(/\D/g, '');
  if (numLimpo.length < 11) return res.status(400).json({ error: 'Telefone inválido' });
  try {
    const notionResp = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN_COMMEDIA}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: NOTION_DB_COMMEDIA }, properties: {
        'Nome Completo':      { title:        [{ text: { content: nomeCompleto } }] },
        'Nome Social':        { rich_text:    [{ text: { content: nomeSocial || '' } }] },
        'CPF':                { rich_text:    [{ text: { content: cpf } }] },
        'RG':                 { rich_text:    [{ text: { content: rg || '' } }] },
        'Telefone':           { phone_number: telefone },
        'Email':              { email:        email },
        'CEP':                { rich_text:    [{ text: { content: cep } }] },
        'Rua':                { rich_text:    [{ text: { content: rua } }] },
        'Número':             { rich_text:    [{ text: { content: numero } }] },
        'Complemento':        { rich_text:    [{ text: { content: complemento || '' } }] },
        'Bairro':             { rich_text:    [{ text: { content: bairro } }] },
        'Cidade':             { rich_text:    [{ text: { content: cidade } }] },
        'Estado':             { rich_text:    [{ text: { content: estado } }] },
        'Forma de Pagamento': { select:       { name: formaPagamento } },
        'Status':             { select:       { name: 'Pendente' } },
        'Assinatura':         { rich_text:    [{ text: { content: assinatura } }] },
        'Observações':        { rich_text:    [{ text: { content: observacoes || '' } }] },
      }}),
    });
    if (!notionResp.ok) { const err = await notionResp.json(); console.error('[commedia]', err); return res.status(500).json({ error: 'Erro Notion' }); }
    const primeiroNome = nomeCompleto.split(' ')[0];
    const isPix = formaPagamento.includes('Pix');
    const msgAdmin = `🎭 *Nova inscrição — Commedia Dell'Arte*\n\n👤 ${nomeCompleto}\n📱 ${telefone}\n📧 ${email}\n💳 ${formaPagamento}`;
    const msgUser = `Olá, ${primeiroNome}! 🎭\n\nSua inscrição no *Workshop de Commedia Dell'Arte* foi recebida!\n\n📅 07, 14, 21 e 28 de agosto\n⏰ Sextas, 10h às 13h\n\n${isPix ? 'Envie R$ 290,00 via Pix para *fabio@cialiquidificador.com.br* e mande o comprovante aqui.' : 'Acesse o link de pagamento parcelado para concluir sua inscrição.'}\n\nQualquer dúvida é só responder aqui! ✨`;
    try { const adminId = await getOrCreateContactId('5511986899433'); if (adminId) await sendDigisacMessage(adminId, msgAdmin); } catch(e) {}
    try { const userId = await getOrCreateContactId('55' + numLimpo); if (userId) await sendDigisacMessage(userId, msgUser); } catch(e) {}
    return res.json({ ok: true });
  } catch (err) { console.error('[commedia]', err.message); return res.status(500).json({ error: 'Erro interno' }); }
});

// ============================================================
// APP DO PRODUTOR — Relatórios de Apresentação
// ============================================================
const { google } = require('googleapis');
const { Readable } = require('stream');

const NOTION_DB_APRESENTACOES = '2b9c45031f7380828d34f47353b066e7';
const GOOGLE_DRIVE_FOLDER_ID = '1LQbLtnDfxrnTZkH50_ce8bLWnIngNv_y';

function getGoogleAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/drive.file'],
  });
}

async function uploadFotoDrive(auth, base64Data, filename) {
  const drive = google.drive({ version: 'v3', auth });
  const mimeType = base64Data.includes('data:image/png') ? 'image/png' : 'image/jpeg';
  const base64Clean = base64Data.replace(/^data:image\/\w+;base64,/, '');
  const buffer = Buffer.from(base64Clean, 'base64');
  const response = await drive.files.create({
    requestBody: { name: filename, parents: [GOOGLE_DRIVE_FOLDER_ID] },
    media: { mimeType, body: Readable.from(buffer) },
    fields: 'id,webViewLink',
  });
  await drive.permissions.create({
    fileId: response.data.id,
    requestBody: { role: 'reader', type: 'anyone' },
  });
  return response.data.webViewLink;
}

function slugify(text) {
  return (text || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9\s]/g, '').trim().replace(/\s+/g, '-').slice(0, 30);
}

app.get('/apresentacoes-hoje', async (req, res) => {
  function hojeBrasilia() {
    const agora = new Date();
    const brasilia = new Date(agora.getTime() - 3 * 60 * 60 * 1000);
    return brasilia.toISOString().split('T')[0];
  }
  const hoje = req.query.data || hojeBrasilia();
  try {
    const r = await fetch(`https://api.notion.com/v1/databases/${NOTION_DB_APRESENTACOES}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: {
          and: [
            { property: 'Data da Apresentação', date: { on_or_before: hoje } },
            { property: 'Data da Apresentação', date: { on_or_after: (() => { const d = new Date(hoje + 'T00:00:00'); d.setDate(d.getDate() - 2); return d.toISOString().split('T')[0]; })() } },
          ],
        },
        sorts: [{ property: 'Data da Apresentação', direction: 'descending' }, { property: 'Horário Apresentação', direction: 'ascending' }],
      }),
    });
    const data = await r.json();
    if (!r.ok) return res.status(500).json({ ok: false, error: data.message });
    async function nomeTituloDaPagina(pageId) {
      try {
        const rp = await fetch('https://api.notion.com/v1/pages/' + pageId, {
          headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
        });
        if (!rp.ok) return '';
        const pageData = await rp.json();
        const props = pageData.properties || {};
        for (const key in props) {
          if (props[key].type === 'title') {
            return (props[key].title?.[0]?.plain_text || '').trim();
          }
        }
        return '';
      } catch (e) { return ''; }
    }

    const apresentacoes = await Promise.all(data.results.map(async page => {
      const p = page.properties;
      const producaoRelation = p['Produção Liqui']?.relation || [];
      const trabalhoRelation = p['🎭 Trabalhos']?.relation || [];

      const produtorNome = producaoRelation.length > 0 ? await nomeTituloDaPagina(producaoRelation[0].id) : '';
      const trabalhoNome  = trabalhoRelation.length  > 0 ? await nomeTituloDaPagina(trabalhoRelation[0].id)  : '';

      return {
        id: page.id,
        trabalho: trabalhoNome,
        local: p['LOCAL']?.title?.[0]?.plain_text || '',
        localNome: p['Local']?.place?.name || p['LOCAL']?.title?.[0]?.plain_text || '',
        localEndereco: p['Local']?.place?.address || '',
        horario: p['Horário Apresentação']?.rich_text?.[0]?.plain_text || '',
        data: p['Data da Apresentação']?.date?.start || hoje,
        produtor: produtorNome,
        jaTemRelatorio: !!(p['Público']?.number),
      };
    }));
    res.json({ ok: true, apresentacoes, data: hoje });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/relatorio-apresentacao', async (req, res) => {
  const { notionPageId, local, localNome, horario, data, produtor, publico, intercorrencia, fotos } = req.body;
  if (!notionPageId || !produtor || publico === undefined) return res.status(400).json({ error: 'Campos obrigatórios faltando' });
  try {
    const msToken = await getMicrosoftToken();
    const localSlug = slugify(localNome || local);
    const produtorSlug = slugify(produtor);
    const fotosValidas = (fotos || []).filter(Boolean);

    const nomePasta = `${data}_${localSlug}_${produtorSlug}`;
    const subpastaId = await criarOuObterSubpasta(msToken, nomePasta);

    const labelsFotos = ['inicio', 'meio', 'fim', 'publico'];
    const uploadPromises = fotosValidas.map((foto, i) => {
      const filename = `${labelsFotos[i] || 'foto' + (i+1)}.jpg`;
      return uploadFotoParaPasta(msToken, subpastaId, foto, filename);
    });
    await Promise.all(uploadPromises);

    const linkPasta = await criarLinkCompartilhamento(msToken, subpastaId);
    const links = [linkPasta];
    const notionResp = await fetch(`https://api.notion.com/v1/pages/${notionPageId}`, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: {
        'Público': { number: parseInt(publico) },
        'Intercorrência': { rich_text: [{ text: { content: intercorrencia || 'Nenhuma' } }] },
        'Fotos Drive': { url: links[0] || null },
      }}),
    });
    if (!notionResp.ok) { const err = await notionResp.json(); return res.status(500).json({ error: 'Erro Notion', detail: err }); }
    res.json({ ok: true, fotos: links });
  } catch (err) {
    console.error('[produtor]', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.get('/teste-notion', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/2b9c45031f7380828d34f47353b066e7/query', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + NOTION_TOKEN,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json',
      },
      body: '{}',
    });
    const data = await r.json();
    res.json({ ok: true, total: data.results ? data.results.length : 0 });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get('/teste-notion2', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/2b9c45031f7380828d34f47353b066e7/query', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + process.env.NOTION_TOKEN_COMMEDIA,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json',
      },
      body: '{}',
    });
    const text = await r.text();
    res.json({ ok: true, status: r.status, preview: text.slice(0, 200) });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get('/teste-http', async (req, res) => {
  try {
    const r = await fetch('https://httpbin.org/get');
    const data = await r.json();
    res.json({ ok: true, origin: data.origin });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

const { Client } = require('@notionhq/client');

app.get('/teste-notion-sdk', async (req, res) => {
  try {
    const notion = new Client({ auth: process.env.NOTION_TOKEN });
    const response = await notion.databases.query({
      database_id: '2b9c45031f7380828d34f47353b066e7',
      page_size: 1,
    });
    res.json({ ok: true, total: response.results.length });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});


// ============================================================
// YOGA — Hatha Yoga com Giulia Hoff
// ============================================================
const PROF_GIULIA = '5512988222584';
const YOGA_LIMITE_EXPERIMENTAL = 10;
const YOGA_LIMITE_REPOSICAO    = 11;

const YOGA_TURMAS = {
  'quarta-7h': { nome: 'Quarta 7h', dia: 'Quarta', horario: '07:00', label: 'quarta-feira as 7h' },
  'quarta-8h': { nome: 'Quarta 8h', dia: 'Quarta', horario: '08:00', label: 'quarta-feira as 8h' },
  'sexta-7h':  { nome: 'Sexta 7h',  dia: 'Sexta',  horario: '07:00', label: 'sexta-feira as 7h'  },
  'sexta-8h':  { nome: 'Sexta 8h',  dia: 'Sexta',  horario: '08:00', label: 'sexta-feira as 8h'  },
};

app.get('/vagas-yoga', async (req, res) => {
  const { turmaId, data, tipo } = req.query;
  const turma = YOGA_TURMAS[turmaId];
  if (!turma) return res.status(400).json({ error: 'Turma invalida' });
  const limite = tipo === 'reposicao' ? YOGA_LIMITE_REPOSICAO : YOGA_LIMITE_EXPERIMENTAL;
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Modalidade', select: { equals: 'Yoga' } }, page_size: 100 }),
    });
    if (!r.ok) { const t = await r.text(); throw new Error('Notion ' + r.status + ': ' + t); }
    const result = await r.json();
    const records = (result.results || []).filter(p => (p.properties?.Turma?.select?.name || '') === turma.nome);
    let count;
    if (tipo === 'reposicao') {
      const ativas = records.filter(p => p.properties?.Status?.select?.name === 'Ativa').length;
      const repos = records.filter(p => { const obs = p.properties?.['Observacoes']?.rich_text?.[0]?.plain_text || p.properties?.['Observacoes']?.text?.[0]?.plain_text || ''; return obs.includes(data) && obs.toLowerCase().includes('repos'); }).length;
      count = ativas + repos;
    } else {
      count = records.filter(p => { const obs = p.properties?.['Observacoes']?.rich_text?.[0]?.plain_text || ''; return obs.includes(data); }).length;
    }
    res.json({ ok: true, vagas: Math.max(0, limite - count), limite, ocupadas: count });
  } catch (err) {
    console.error('[yoga/vagas]', err.message);
    res.json({ ok: false, vagas: limite, limite, error: err.message });
  }
});

app.post('/inscricao-yoga', async (req, res) => {
  const { nome, whatsapp, email, turmaId, data, tipo } = req.body;
  if (!nome || !whatsapp || !turmaId || !data || !tipo) return res.status(400).json({ error: 'Campos obrigatorios faltando' });
  const turma = YOGA_TURMAS[turmaId];
  if (!turma) return res.status(400).json({ error: 'Turma invalida' });
  const numLimpo = whatsapp.replace(/\D/g, '');
  if (numLimpo.length < 11) return res.status(400).json({ error: 'WhatsApp invalido' });
  const tipoTexto = tipo === 'reposicao' ? 'Reposicao' : 'Aula experimental gratuita';
  const horDisplay = turma.horario.replace(':00', 'h');
  const dataFmt = data.split('-').reverse().join('/');
  const primeiroNome = nome.split(' ')[0];
  const obs = data + ' - ' + tipoTexto + ' - ' + turma.label + (email ? ' - ' + email : '');
  try {
    const notionResp = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: {
        'Nome':        { title:        [{ text: { content: nome } }] },
        'Contato':     { phone_number: whatsapp },
        'Turma':       { select:       { name: turma.nome } },
        'Modalidade':  { select:       { name: 'Yoga' } },
        'Status':      { select:       { name: 'Experimental' } },
        'Professor':   { select:       { name: 'Giulia' } },
        'Dia':         { select:       { name: turma.dia } },
        'Observações': { rich_text:    [{ text: { content: obs } }] },
      }}),
    });
    if (!notionResp.ok) { const e = await notionResp.json(); console.error('[yoga/notion]', e); return res.status(500).json({ error: 'Erro ao salvar', detail: e }); }
    const numBr = '55' + numLimpo;
    const msgAluna = 'Ola, ' + primeiroNome + '! Sua vaga na aula de *Hatha Yoga* com Giulia Hoff esta confirmada!\n\nData: ' + dataFmt + '\nHorario: ' + horDisplay + '\nEspaco Liquidificador - Rua Dr. Carvalho de Mendonca, 67, Campos Eliseos\n\nQualquer duvida e so responder aqui!';
    const msgGiulia = 'Nova inscricao - Yoga\n\nNome: ' + nome + '\nWhatsApp: ' + whatsapp + (email ? '\nEmail: ' + email : '') + '\nData: ' + dataFmt + ' - ' + horDisplay + '\nTipo: ' + tipoTexto;
    const msgAdmin  = 'Nova inscricao - Yoga\n\nNome: ' + nome + '\nWhatsApp: ' + whatsapp + '\nData: ' + dataFmt + ' - ' + turma.label + '\nTipo: ' + tipoTexto;
    const dataAula = new Date(data + 'T' + turma.horario + ':00-03:00');
    const lembrete = new Date(dataAula); lembrete.setDate(lembrete.getDate() - 1); lembrete.setHours(8, 0, 0, 0);
    const msgLembrete = 'Ola, ' + primeiroNome + '! Lembrando que amanha voce tem aula de *Hatha Yoga* com Giulia Hoff!\n\nHorario: ' + horDisplay + '\nRua Dr. Carvalho de Mendonca, 67 - Campos Eliseos';
    try {
      const contactId = await getOrCreateContactId(numBr);
      if (contactId) {
        await fetch(DIGISAC_BASE + '/messages', { method: 'POST', headers: digisacHeaders, body: JSON.stringify({ text: msgAluna, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot' }) });
        await fetch(DIGISAC_BASE + '/messages', { method: 'POST', headers: digisacHeaders, body: JSON.stringify({ text: msgLembrete, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot', scheduledAt: lembrete.toISOString() }) });
      }
    } catch(e) { console.error('[yoga/msg-aluna]', e.message); }
    try { await enviarWhatsAppComHorarioComercial(PROF_GIULIA, msgGiulia); } catch(e) { console.error('[yoga/msg-giulia]', e.message); }
    try { await enviarWhatsApp('5511986899433', msgAdmin); } catch(e) { console.error('[yoga/msg-admin]', e.message); }
    res.json({ ok: true });
  } catch (err) {
    console.error('[yoga/inscricao]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ============================================================
// APP DE PRESENÇA — Registro de presença por professor/turma
// ============================================================
const PRESENCAS_DB = '282c0bc0-06d8-4828-acae-d5ac35388318';
const WHATSAPP_FABIO = '5511989946586';
const WHATSAPP_CIA = '5511986899433';

const PROFESSORES_PRESENCA = [
  {
    nome: 'André / Renata',
    turmas: [
      { turma: 'Segunda 10h', dia: 'Segunda', horario: '10:00', modalidade: 'Circo - Acrobacia' },
    ],
  },
  {
    nome: 'Gabi',
    turmas: [
      { turma: 'Segunda 18h', dia: 'Segunda', horario: '18:00', modalidade: 'Aéreos' },
      { turma: 'Segunda 19h', dia: 'Segunda', horario: '19:00', modalidade: 'Aéreos' },
    ],
  },
  {
    nome: 'Giulia',
    turmas: [
      { turma: 'Quarta 7h', dia: 'Quarta', horario: '07:00', modalidade: 'Yoga' },
      { turma: 'Quarta 8h', dia: 'Quarta', horario: '08:00', modalidade: 'Yoga' },
      { turma: 'Sexta 7h', dia: 'Sexta', horario: '07:00', modalidade: 'Yoga' },
      { turma: 'Sexta 8h', dia: 'Sexta', horario: '08:00', modalidade: 'Yoga' },
    ],
  },
  {
    nome: 'Guilherme',
    turmas: [
      { turma: 'Quinta 8h', dia: 'Quinta', horario: '08:00', modalidade: 'Aéreos' },
    ],
  },
  {
    nome: 'Gustra',
    turmas: [
      { turma: 'Quarta 18h', dia: 'Quarta', horario: '18:00', modalidade: 'Aéreos' },
      { turma: 'Quarta 19h', dia: 'Quarta', horario: '19:00', modalidade: 'Aéreos' },
    ],
  },
  {
    nome: 'Talita',
    turmas: [
      { turma: 'Terça 8h', dia: 'Terça', horario: '08:00', modalidade: 'Aéreos' },
      { turma: 'Terça 9h', dia: 'Terça', horario: '09:00', modalidade: 'Aéreos' },
    ],
  },
  {
    nome: 'Titzi',
    turmas: [
      { turma: 'Terça 18h', dia: 'Terça', horario: '18:00', modalidade: 'Circo Infantil' },
      { turma: 'Quarta 9h30', dia: 'Quarta', horario: '09:30', modalidade: 'Circo Infantil' },
    ],
  },
];

// GET /professores-presenca
app.get('/professores-presenca', (req, res) => {
  res.json({ ok: true, professores: PROFESSORES_PRESENCA });
});

// GET /alunas-turma?turma=Segunda 18h&modalidade=Aéreos&data=2026-07-06
app.get('/alunas-turma', async (req, res) => {
  const { turma, modalidade, data } = req.query;
  if (!turma || !modalidade) return res.status(400).json({ error: 'Turma e modalidade obrigatórios' });

  try {
    const rAtivas = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: {
          and: [
            { property: 'Turma', select: { equals: turma } },
            { property: 'Modalidade', select: { equals: modalidade } },
            { property: 'Status', select: { equals: 'Ativa' } },
          ],
        },
        page_size: 100,
      }),
    });
    const dataAtivas = await rAtivas.json();
    const ativas = (dataAtivas.results || [])
      .map(p => ({
        id: p.id,
        nome: p.properties?.Nome?.title?.[0]?.plain_text || '',
        tipo: 'Regular',
      }))
      .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));

    let experimentais = [];
    if (data) {
      const rExp = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: {
            and: [
              { property: 'Turma', select: { equals: turma } },
              { property: 'Modalidade', select: { equals: modalidade } },
              { property: 'Status', select: { equals: 'Experimental' } },
            ],
          },
          page_size: 100,
        }),
      });
      const dataExp = await rExp.json();
      experimentais = (dataExp.results || [])
        .filter(p => {
          const obs = p.properties?.Observações?.rich_text?.[0]?.plain_text || '';
          return obs.includes(data);
        })
        .map(p => ({
          id: p.id,
          nome: p.properties?.Nome?.title?.[0]?.plain_text || '',
          contato: p.properties?.Contato?.phone_number || '',
          tipo: 'Experimental',
        }))
        .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
    }

    res.json({ ok: true, ativas, experimentais });
  } catch (err) {
    console.error('[presenca] alunas-turma:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /registrar-presenca
app.post('/registrar-presenca', async (req, res) => {
  const { professor, turma, modalidade, data, presencas } = req.body;

  if (!professor || !turma || !data || !Array.isArray(presencas)) {
    return res.status(400).json({ error: 'Campos obrigatórios: professor, turma, data, presencas' });
  }

  try {
    const criados = [];

    // Buscar registros ja existentes hoje para essa turma (evitar duplicatas)
    const rExistentes = await fetch('https://api.notion.com/v1/databases/' + PRESENCAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Turma', select: { equals: turma } },
          { property: 'Data', date: { equals: data } },
        ]},
        page_size: 100,
      }),
    });
    const dExistentes = await rExistentes.json();
    const mapaExistentes = {};
    (dExistentes.results || []).forEach(page => {
      const alunaId = page.properties?.Aluna?.relation?.[0]?.id;
      if (alunaId) mapaExistentes[alunaId] = page.id;
    });

    for (const p of presencas) {
      const pageIdExistente = mapaExistentes[p.alunaId];
      const method = pageIdExistente ? 'PATCH' : 'POST';
      const url = pageIdExistente ? ('https://api.notion.com/v1/pages/' + pageIdExistente) : 'https://api.notion.com/v1/pages';
      const bodyObj = pageIdExistente
        ? { properties: {
            'Status': { select: { name: p.status } },
            'Tipo': { select: { name: p.tipo || 'Regular' } },
          }}
        : { parent: { database_id: PRESENCAS_DB }, properties: {
            'Nome': { title: [{ text: { content: `${p.nome} - ${data}` } }] },
            'Aluna': { relation: [{ id: p.alunaId }] },
            'Turma': { select: { name: turma } },
            'Professor': { select: { name: professor } },
            'Data': { date: { start: data } },
            'Status': { select: { name: p.status } },
            'Tipo': { select: { name: p.tipo || 'Regular' } },
          }};

      const notionResp = await fetch(url, {
        method,
        headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify(bodyObj),
      });

      if (notionResp.ok) {
        const created = await notionResp.json();
        criados.push({ ...p, pageId: created.id });
      } else {
        const errData = await notionResp.json();
        console.error('[presenca] erro ao criar registro:', errData);
      }
    }

    const experimentaisRegistradas = presencas.filter(p => p.tipo === 'Experimental');
    for (const exp of experimentaisRegistradas) {
      const resultado = exp.status === 'Presente' ? 'COMPARECEU ✅' : 'FALTOU ❌';
      const msg = `🎪 *Resultado Aula Experimental*\n\nAluna: ${exp.nome}\nTurma: ${turma}\nProfessor: ${professor}\nData: ${data.split('-').reverse().join('/')}\nResultado: ${resultado}`;
      try { await enviarWhatsApp(WHATSAPP_FABIO, msg); } catch(e) { console.error('[presenca] wpp fabio exp:', e.message); }
      try { await enviarWhatsApp(WHATSAPP_CIA, msg); } catch(e) { console.error('[presenca] wpp cia exp:', e.message); }

      // Mensagem direta para a aluna (chatbot de acompanhamento)
      if (exp.contato) {
        const numeroAluna = normalizarTelefone(exp.contato);
        if (numeroAluna) {
          try {
            if (exp.status === 'Presente') {
              await enviarWhatsAppComHorarioComercial(numeroAluna, msgComparecimento(exp.nome, modalidade));
            } else {
              await enviarWhatsAppComHorarioComercial(numeroAluna, msgFalta(exp.nome, modalidade));
              CONVERSAS_ESTADO[numeroAluna] = { estado: 'aguardando_sim_nao', alunaId: exp.alunaId, nome: exp.nome, modalidade };
            }
          } catch(e) { console.error('[presenca] wpp aluna exp:', e.message); }
        }
      }
    }

    // Convite de reposicao para quem faltou (exceto Acro, que so tem 1 horario)
    const MAPA_LINKS_REPOSICAO = {
      'Aéreos': 'https://agende-aereos.ciadoliquidificador.com.br',
      'Circo Infantil': 'https://agende-infantil.ciadoliquidificador.com.br',
      'Yoga': 'https://agende-yoga.ciadoliquidificador.com.br',
    };

    const faltasParaConvite = presencas.filter(p => p.tipo === 'Regular' && p.status === 'Falta');
    const linkReposicao = MAPA_LINKS_REPOSICAO[modalidade];

    if (linkReposicao) {
      for (const falta of faltasParaConvite) {
        try {
          const alunaResp = await fetch('https://api.notion.com/v1/pages/' + falta.alunaId, {
            headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
          });
          if (alunaResp.ok) {
            const alunaData = await alunaResp.json();
            const telefone = alunaData.properties?.Contato?.phone_number || '';
            if (telefone) {
              const primeiroNome = falta.nome.split(' ')[0];
              const msgConvite = 'Oi, ' + primeiroNome + '! 🌿\n\nSentimos sua falta na aula de ' + modalidade + '! Está tudo bem? 💛\n\nSe quiser, você pode repor a aula em outro horário disponível:\n' + linkReposicao + '\n\nQualquer coisa, é só chamar aqui. Um abraço! 🤗';
              const numLimpo = telefone.replace(/\D/g, '');
              const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
              await enviarWhatsAppComHorarioComercial(numBr, msgConvite);
            }
          }
        } catch (e) {
          console.error('[presenca] erro ao enviar convite reposicao:', e.message);
        }
      }
    }

    const faltasRegulares = presencas.filter(p => p.tipo === 'Regular' && p.status === 'Falta');
    for (const falta of faltasRegulares) {
      try {
        const rHist = await fetch(`https://api.notion.com/v1/databases/${PRESENCAS_DB}/query`, {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
          body: JSON.stringify({
            filter: {
              and: [
                { property: 'Aluna', relation: { contains: falta.alunaId } },
                { property: 'Tipo', select: { equals: 'Regular' } },
              ],
            },
            sorts: [{ property: 'Data', direction: 'descending' }],
            page_size: 20,
          }),
        });
        const histData = await rHist.json();
        const registros = (histData.results || []).map(p => ({
          data: p.properties?.Data?.date?.start || '',
          status: p.properties?.Status?.select?.name || '',
        }));

        const doisUltimos = registros.slice(0, 2);
        const consecutivas = doisUltimos.length === 2 && doisUltimos.every(r => r.status === 'Falta');

        const mesAtual = data.slice(0, 7);
        const faltasNoMes = registros.filter(r => r.status === 'Falta' && r.data.startsWith(mesAtual)).length;

        if (consecutivas || faltasNoMes >= 2) {
          const motivo = consecutivas ? '2 faltas seguidas' : `${faltasNoMes} faltas no mês`;
          const msgAlerta = `⚠️ *Alerta de Faltas*\n\nAluna: ${falta.nome}\nTurma: ${turma}\nProfessor: ${professor}\nMotivo: ${motivo}`;
          try { await enviarWhatsApp(WHATSAPP_FABIO, msgAlerta); } catch(e) { console.error('[presenca] wpp fabio falta:', e.message); }
          try { await enviarWhatsApp(WHATSAPP_CIA, msgAlerta); } catch(e) { console.error('[presenca] wpp cia falta:', e.message); }
        }
      } catch (e) {
        console.error('[presenca] erro ao checar historico:', e.message);
      }
    }

    res.json({ ok: true, criados: criados.length });
  } catch (err) {
    console.error('[presenca] registrar-presenca:', err.message);
    res.status(500).json({ error: err.message });
  }
});


// ============================================================
// CHATBOT — Acompanhamento de aulas experimentais via WhatsApp
// ============================================================
const CONVERSAS_ESTADO = {}; // { numero: { estado, alunaId, nome, modalidade } }

const LINKS_AGENDAMENTO = {
  'Aéreos': 'https://agende-aereos.ciadoliquidificador.com.br',
  'Circo Infantil': 'https://agende-infantil.ciadoliquidificador.com.br',
  'Yoga': 'https://agende-yoga.ciadoliquidificador.com.br',
  'Circo - Acrobacia': 'https://agende-acro.ciadoliquidificador.com.br',
};

function normalizarTelefone(tel) {
  const digitos = (tel || '').replace(/\D/g, '');
  if (digitos.length === 11) return '55' + digitos;
  if (digitos.length === 13 && digitos.startsWith('55')) return digitos;
  return digitos;
}

function interpretarSimNao(texto) {
  const t = (texto || '').trim().toLowerCase();
  const simPalavras = ['sim', 's', 'quero', 'yes', 'claro', 'com certeza', 'pode', 'ok', 'positivo'];
  const naoPalavras = ['não', 'nao', 'n', 'no', 'negativo'];
  if (simPalavras.some(p => t === p || t.startsWith(p + ' '))) return 'sim';
  if (naoPalavras.some(p => t === p || t.startsWith(p + ' '))) return 'nao';
  return null;
}

function interpretarMotivo(texto) {
  const t = (texto || '').trim().toLowerCase();
  if (t === '1' || t.includes('preço') || t.includes('preco') || t.includes('caro') || t.includes('valor')) return 'Preço';
  if (t === '2' || t.includes('local') || t.includes('longe') || t.includes('distanc')) return 'Localização';
  if (t === '3' || t.includes('horár') || t.includes('horar') || t.includes('hora')) return 'Horário';
  if (t === '4' || t.includes('outro')) return 'Outro';
  return null;
}

function msgFalta(nome, modalidade) {
  const primeiroNome = nome.split(' ')[0];
  return `Oi, ${primeiroNome}! 🌿\n\nNotamos que você não conseguiu vir na aula experimental de ${modalidade}. Tudo bem?\n\nQuer remarcar para outro dia? Responde *SIM* ou *NÃO* por aqui! 😊`;
}

function msgComparecimento(nome, modalidade) {
  const primeiroNome = nome.split(' ')[0];
  return `Oi, ${primeiroNome}! 🎪\n\nQue bom que você veio conhecer a aula de ${modalidade}! Como foi a experiência? 💛\n\nFicou com vontade de continuar com a gente? Se sim, é só responder aqui que te ajudamos com a matrícula!`;
}

function msgLinkRemarcar(modalidade) {
  const link = LINKS_AGENDAMENTO[modalidade] || 'https://reposicao.ciadoliquidificador.com.br';
  return `Ótimo! Aqui está o link para escolher um novo horário:\n${link}`;
}

function msgMotivo() {
  return `Sem problemas! Pra gente entender melhor e sempre melhorar, qual foi o motivo?\n\n1️⃣ Preço\n2️⃣ Localização\n3️⃣ Horário\n4️⃣ Outro motivo`;
}

function msgAgradecimentoMotivo() {
  return `Obrigado por nos contar! 💛 Se um dia quiser voltar, estaremos de portas abertas.`;
}

function msgNaoEntendiSimNao() {
  return `Desculpa, não entendi sua resposta 🙏\n\nVocê quer remarcar a aula experimental? Responde *SIM* ou *NÃO*.`;
}

function msgNaoEntendiMotivo() {
  return `Desculpa, não entendi 🙏\n\nEscolha uma opção:\n\n1️⃣ Preço\n2️⃣ Localização\n3️⃣ Horário\n4️⃣ Outro motivo`;
}

async function salvarMotivoDesistencia(alunaId, motivo) {
  try {
    await fetch(`https://api.notion.com/v1/pages/${alunaId}`, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Motivo Desistência': { select: { name: motivo } },
        },
      }),
    });
  } catch (e) {
    console.error('[chatbot] erro ao salvar motivo:', e.message);
  }
}

// POST /webhook-digisac — recebe mensagens recebidas via Digisac
app.post('/webhook-digisac', async (req, res) => {
  res.status(200).json({ ok: true }); // responde rápido, processa depois

  try {
    const body = req.body || {};

    const isFromMe = body?.data?.isFromMe || body?.data?.fromMe || false;
    if (isFromMe) return; // ignora mensagens enviadas por nós mesmos

    const tipoMsg = body?.data?.type;
    if (tipoMsg !== 'chat' && tipoMsg !== 'image' && tipoMsg !== 'document') return; // ignora reacoes, videos, tickets, etc.

    console.log('[webhook-digisac] payload completo:', JSON.stringify(body).slice(0, 2000));

    const texto = body?.data?.text || body?.data?.body || body?.data?.message || body?.data?.data?.text || '';
    const contactId = body?.data?.contactId || '';
    const ticketId = body?.data?.ticketId || '';

    let numero = '';
    if (contactId) {
      try {
        const contactResp = await fetch(DIGISAC_BASE + '/contacts/' + contactId, {
          headers: { 'Authorization': 'Bearer ' + DIGISAC_TOKEN },
        });
        if (contactResp.ok) {
          const contactData = await contactResp.json();
          const numRaw = contactData?.data?.number || contactData?.number || '';
          numero = normalizarTelefone(numRaw);
          console.log('[webhook-digisac] contactId->numero:', contactId, '->', numero);
        } else {
          console.log('[webhook-digisac] erro ao buscar contato, status:', contactResp.status);
        }
      } catch (e) {
        console.log('[webhook-digisac] erro ao buscar contato:', e.message);
      }
    }

    if (!numero || (!texto && tipoMsg === 'chat')) {
      console.log('[webhook-digisac] numero ou texto vazio, ignorando. numero=', numero, 'texto=', texto);
      return;
    }

    const estado = CONVERSAS_ESTADO[numero];
    if (!estado) {
      console.log('[webhook-digisac] sem estado de conversa para', numero, '— ignorando');
      return;
    }

    if (estado.estado === 'aguardando_confirmacao_ensaio' || estado.estado === 'aguardando_comprovante_imagem') {
      await processarRespostaSalaEnsaio(numero, texto, tipoMsg, ticketId, estado);
      return;
    }

    if (estado.estado === 'aguardando_sim_nao') {
      const resposta = interpretarSimNao(texto);
      if (resposta === 'sim') {
        await enviarWhatsApp(numero, msgLinkRemarcar(estado.modalidade));
        delete CONVERSAS_ESTADO[numero];
      } else if (resposta === 'nao') {
        await enviarWhatsApp(numero, msgMotivo());
        CONVERSAS_ESTADO[numero] = { ...estado, estado: 'aguardando_motivo' };
      } else {
        await enviarWhatsApp(numero, msgNaoEntendiSimNao());
      }
      return;
    }

    if (estado.estado === 'aguardando_motivo') {
      const motivo = interpretarMotivo(texto);
      if (motivo) {
        await salvarMotivoDesistencia(estado.alunaId, motivo);
        await enviarWhatsApp(numero, msgAgradecimentoMotivo());
        delete CONVERSAS_ESTADO[numero];
      } else {
        await enviarWhatsApp(numero, msgNaoEntendiMotivo());
      }
      return;
    }
  } catch (err) {
    console.error('[webhook-digisac] erro:', err.message);
  }
});

// GET /presencas-do-dia?turma=X&data=YYYY-MM-DD
app.get('/presencas-do-dia', async (req, res) => {
  const { turma, data } = req.query;
  if (!turma || !data) return res.status(400).json({ error: 'Turma e data obrigatorios' });
  try {
    const r = await fetch(`https://api.notion.com/v1/databases/${PRESENCAS_DB}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: {
          and: [
            { property: 'Turma', select: { equals: turma } },
            { property: 'Data', date: { equals: data } },
          ],
        },
        page_size: 100,
      }),
    });
    const d = await r.json();
    const registros = (d.results || []).map(p => ({
      pageId: p.id,
      alunaId: p.properties?.Aluna?.relation?.[0]?.id || null,
      status: p.properties?.Status?.select?.name || null,
    })).filter(x => x.alunaId);
    res.json({ ok: true, registros });
  } catch (err) {
    console.error('[presenca] presencas-do-dia:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ============================================================
// SALA DE ENSAIO — Agendamento com calculo automatico de valor
// ============================================================
const SALA_ENSAIO_DB = 'd3a426a5-04c9-4fba-a77f-790d00468cb7';
const VALOR_ENSAIO_DB = '226c4503-1f73-8066-95d5-f03ecb39f7f5';
const SUPERVISOR_FDS = 37.50;
const CALENDARIO_FERIADOS_ID = 'en.brazilian#holiday@group.v.calendar.google.com';

// Cache de precos (5 minutos)
let _cachePrecos = null;
let _cachePrecosTimestamp = 0;

async function getTabelaPrecos() {
  const agora = Date.now();
  if (_cachePrecos && (agora - _cachePrecosTimestamp) < 5 * 60 * 1000) {
    return _cachePrecos;
  }
  const r = await fetch('https://api.notion.com/v1/databases/' + VALOR_ENSAIO_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ page_size: 100 }),
  });
  const d = await r.json();
  const tabela = {};
  (d.results || []).forEach(page => {
    const qtdStr = page.properties?.['Quantidade de Pessoas']?.title?.[0]?.plain_text || '';
    const qtd = parseInt(qtdStr, 10);
    const valor = page.properties?.['Valor Total por Hora (R$)']?.number;
    if (!isNaN(qtd) && valor != null) tabela[qtd] = valor;
  });
  _cachePrecos = tabela;
  _cachePrecosTimestamp = agora;
  return tabela;
}

async function getValorHoraDiaUtil(qtd) {
  const tabela = await getTabelaPrecos();
  const qtds = Object.keys(tabela).map(Number).sort((a,b) => a-b);
  if (qtds.length === 0) return 55; // fallback de seguranca
  const qtdAlvo = Math.max(1, Math.round(qtd));
  let escolhido = qtds[0];
  for (const q of qtds) { if (q <= qtdAlvo) escolhido = q; }
  return tabela[escolhido];
}
async function getValorHoraFDS(qtd) {
  return (await getValorHoraDiaUtil(qtd)) + SUPERVISOR_FDS;
}

function getDescontoPacote(totalHoras, ehFDS) {
  if (ehFDS) {
    if (totalHoras >= 40) return 0.15;
    if (totalHoras >= 20) return 0.12;
    if (totalHoras >= 10) return 0.10;
    if (totalHoras >= 5) return 0.05;
    return 0;
  }
  if (totalHoras >= 40) return 0.20;
  if (totalHoras >= 20) return 0.15;
  if (totalHoras >= 10) return 0.10;
  if (totalHoras >= 5) return 0.05;
  return 0;
}

// Cache de feriados por ano
const _cacheFeriados = {};

async function getFeriadosDoAno(ano) {
  if (_cacheFeriados[ano]) return _cacheFeriados[ano];
  try {
    const auth = new (require('googleapis').google.auth.GoogleAuth)({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/calendar.readonly'],
    });
    const authClient = await auth.getClient();
    const calendar = require('googleapis').google.calendar({ version: 'v3', auth: authClient });
    const resp = await calendar.events.list({
      calendarId: CALENDARIO_FERIADOS_ID,
      timeMin: ano + '-01-01T00:00:00Z',
      timeMax: ano + '-12-31T23:59:59Z',
      singleEvents: true,
    });
    const datas = new Set((resp.data.items || []).map(ev => ev.start?.date).filter(Boolean));
    _cacheFeriados[ano] = datas;
    return datas;
  } catch (e) {
    console.error('[sala-ensaio] erro ao buscar feriados:', e.message);
    return new Set(); // fallback: sem feriados detectados
  }
}

async function ehFimDeSemanaOuFeriado(dataStr) {
  const d = new Date(dataStr + 'T12:00:00-03:00');
  const diaSemana = d.getDay();
  if (diaSemana === 0 || diaSemana === 6) return true;
  const ano = dataStr.split('-')[0];
  const feriados = await getFeriadosDoAno(ano);
  return feriados.has(dataStr);
}

function formatarHorasBR(decimal) {
  const h = Math.floor(decimal);
  const m = Math.round((decimal - h) * 60);
  return m > 0 ? h + 'h' + String(m).padStart(2, '0') : h + 'h';
}

function horasEntre(inicio, fim) {
  const [hI, mI] = inicio.split(':').map(Number);
  const [hF, mF] = fim.split(':').map(Number);
  return (hF * 60 + mF - (hI * 60 + mI)) / 60;
}

async function calcularValorEnsaio(qtdPessoas, blocos) {
  let totalHorasDiaUtil = 0;
  let totalHorasFDS = 0;
  const detalhes = [];

  for (const b of blocos) {
    const horas = horasEntre(b.inicio, b.fim);
    const fds = await ehFimDeSemanaOuFeriado(b.data);
    if (fds && horas < 4) {
      return { erro: 'O dia ' + b.data + ' é fim de semana/feriado e precisa de no mínimo 4 horas.' };
    }
    if (fds) totalHorasFDS += horas; else totalHorasDiaUtil += horas;
    detalhes.push({ ...b, horas, fds });
  }

  const descontoDU = getDescontoPacote(totalHorasDiaUtil, false);
  const descontoFDS = getDescontoPacote(totalHorasFDS, true);
  const valorHoraDU = await getValorHoraDiaUtil(qtdPessoas);
  const valorHoraFDS = await getValorHoraFDS(qtdPessoas);
  const valorDU = totalHorasDiaUtil * valorHoraDU * (1 - descontoDU);
  const valorFDS = totalHorasFDS * valorHoraFDS * (1 - descontoFDS);
  const subtotal = valorDU + valorFDS;

  const totalHoras = totalHorasDiaUtil + totalHorasFDS;
  const valorBruto = (totalHorasDiaUtil * valorHoraDU) + (totalHorasFDS * valorHoraFDS);
  const valorPorHoraMedio = totalHoras > 0 ? valorBruto / totalHoras : 0;
  const descontoValor = valorBruto - subtotal;

  return {
    ok: true,
    detalhes,
    totalHorasDiaUtil,
    totalHorasFDS,
    totalHoras: Math.round(totalHoras * 100) / 100,
    valorPorHoraMedio: Math.round(valorPorHoraMedio * 100) / 100,
    valorBruto: Math.round(valorBruto * 100) / 100,
    descontoValor: Math.round(descontoValor * 100) / 100,
    descontoDU: descontoDU * 100,
    descontoFDS: descontoFDS * 100,
    subtotal: Math.round(subtotal * 100) / 100,
  };
}

app.post('/calcular-ensaio', async (req, res) => {
  const { qtdPessoas, blocos, notaFiscal } = req.body;
  if (!qtdPessoas || !Array.isArray(blocos) || blocos.length === 0) {
    return res.status(400).json({ error: 'qtdPessoas e blocos sao obrigatorios' });
  }
  try {
    const resultado = await calcularValorEnsaio(qtdPessoas, blocos);
    if (resultado.erro) return res.json({ ok: false, erro: resultado.erro });

    const subtotal = resultado.subtotal;
    const totalComNota = notaFiscal ? Math.round(subtotal * 1.10 * 100) / 100 : subtotal;
    const deposito = Math.round(totalComNota * 0.5 * 100) / 100;

    res.json({ ok: true, ...resultado, totalFinal: totalComNota, deposito });
  } catch (err) {
    console.error('[sala-ensaio] calcular-ensaio:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/verificar-conflito', async (req, res) => {
  const { data, inicio, fim } = req.body;
  if (!data || !inicio || !fim) return res.status(400).json({ error: 'data, inicio, fim obrigatorios' });

  try {
    const inicioISO = data + 'T' + inicio + ':00-03:00';
    const fimISO = data + 'T' + fim + ':00-03:00';
    const resultado = await verificarDisponibilidade(inicioISO, fimISO);
    res.json({ ok: true, disponivel: resultado.disponivel });
  } catch (err) {
    console.error('[sala-ensaio] verificar-conflito:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/reservar-sala', async (req, res) => {
  const { projeto, coletivo, diretor, contatoNome, whatsapp, qtdPessoas, notaFiscal, tipoEnsaio, blocos, cpfCnpj } = req.body;

  if (!projeto || !whatsapp || !qtdPessoas || !Array.isArray(blocos) || blocos.length === 0) {
    return res.status(400).json({ error: 'Campos obrigatorios faltando' });
  }

  try {
    const resultado = await calcularValorEnsaio(qtdPessoas, blocos);
    if (resultado.erro) return res.json({ ok: false, erro: resultado.erro });

    const subtotal = resultado.subtotal;
    const totalFinal = notaFiscal ? Math.round(subtotal * 1.10 * 100) / 100 : subtotal;
    const deposito = Math.round(totalFinal * 0.5 * 100) / 100;

    const blocosResidentes = [];
    for (const b of blocos) {
      const inicioISO = b.data + 'T' + b.inicio + ':00-03:00';
      const fimISO = b.data + 'T' + b.fim + ':00-03:00';
      const resultadoDisp = await verificarDisponibilidade(inicioISO, fimISO);
      if (!resultadoDisp.disponivel) {
        const motivo = resultadoDisp.motivoResidente
          ? 'O dia ' + b.data + ' das ' + b.inicio + ' às ' + b.fim + ' pertence à residência da Cia Plá e o limite de uso este mês/semana já foi atingido. Escolha outro horário.'
          : 'O dia ' + b.data + ' das ' + b.inicio + ' às ' + b.fim + ' já está reservado. Escolha outro horário.';
        return res.json({ ok: false, erro: motivo });
      }
      if (resultadoDisp.ehResidente) blocosResidentes.push(b);
    }

    const reservaId = 'ENS-' + Date.now();
    const numLimpo = (whatsapp || '').replace(/\D/g, '');
    const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;

    for (const b of blocos) {
      const inicioISO = b.data + 'T' + b.inicio + ':00-03:00';
      const fimISO = b.data + 'T' + b.fim + ':00-03:00';

      const googleEventId = await criarEventoEnsaioExterno({
        projeto, coletivo, diretor, whatsapp, tipoEnsaio, inicioISO, fimISO, reservaId,
      });

      await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: SALA_ENSAIO_DB },
          properties: {
            'Título': { title: [{ text: { content: projeto + ' — ' + b.data + ' (' + b.inicio + '-' + b.fim + ')' } }] },
            'Nome do Projeto': { rich_text: [{ text: { content: projeto } }] },
            'Coletivo / Grupo': { rich_text: [{ text: { content: coletivo || '' } }] },
            'Responsável': { rich_text: [{ text: { content: diretor || '' } }] },
            'Contato': { rich_text: [{ text: { content: contatoNome || '' } }] },
            'WhatsApp': { rich_text: [{ text: { content: whatsapp } }] },
            'Início': { date: { start: inicioISO } },
            'Fim': { date: { start: fimISO } },
            'Quantidade de Pessoas': { number: Number(qtdPessoas) },
            'Tipo de Ensaio': { select: { name: tipoEnsaio || 'Outros' } },
            'Status': { select: { name: 'Pendente' } },
            'Nota Fiscal': { select: { name: notaFiscal ? 'Sim' : 'Não' } },
            'Valor Total': { number: totalFinal },
            'Reserva ID': { rich_text: [{ text: { content: reservaId } }] },
            'Google Event ID': { rich_text: [{ text: { content: googleEventId || '' } }] },
            'CPF/CNPJ Nota Fiscal': { rich_text: [{ text: { content: cpfCnpj || 'A informar' } }] },
          },
        }),
      });
    }

    const resumoDias = blocos.map(b => {
      const dataFmt = b.data.split('-').reverse().join('/');
      return '📅 ' + dataFmt + ' — ' + b.inicio + ' às ' + b.fim;
    }).join('\n');

    const totalHorasMsg = formatarHorasBR(resultado.totalHoras);
    const descontoTexto = resultado.descontoValor > 0 ? '- R$ ' + resultado.descontoValor.toFixed(2) : 'R$ 0,00';
    const valorNotaFiscal = totalFinal - resultado.subtotal;
    const notaTexto = notaFiscal ? '\n📄 Nota fiscal: R$ ' + valorNotaFiscal.toFixed(2) : '';
    const chavePix = notaFiscal ? 'financeiro@cialiquidificador.com.br' : 'fabio@cialiquidificador.com.br';

    const primeiroNome = (contatoNome || projeto).split(' ')[0];
    const msgCliente = 'Olá, ' + primeiroNome + '! 🎭\n\nSeu ensaio está pré-agendado!\n\n🎬 Projeto: ' + projeto + '\n' + resumoDias + '\n\n⏱️ Total de horas: ' + totalHorasMsg + '\n💲 Valor/hora: R$ ' + resultado.valorPorHoraMedio.toFixed(2) + '\n📊 Sub-total: R$ ' + resultado.valorBruto.toFixed(2) + '\n🏷️ Desconto: ' + descontoTexto + '\n📋 Subtotal com desconto: R$ ' + resultado.subtotal.toFixed(2) + notaTexto + '\n💰 Total: R$ ' + totalFinal.toFixed(2) + '\n💵 Sinal para garantir a reserva: R$ ' + deposito.toFixed(2) + '\n🔑 Chave PIX: ' + chavePix + '\n\nPara confirmar, faça o pagamento do sinal e mande o comprovante aqui mesmo.\n\nResponda por aqui:\n1️⃣ Confirmar agendamento\n2️⃣ Cancelar agendamento\n3️⃣ Falar com atendente';

    await enviarWhatsApp(numBr, msgCliente);
    CONVERSAS_ESTADO[numBr] = { estado: 'aguardando_confirmacao_ensaio', reservaId, nome: primeiroNome, valorTotal: totalFinal, deposito, criadoEm: Date.now(), lembreteEnviado: false };

    // Notificar residente sobre os blocos que tomamos do horario dele (ja verificados acima)
    try {
      if (blocosResidentes.length > 0) {
        await notificarResidenteSobreposicao(blocosResidentes);
      }
    } catch (e) {
      console.error('[residentes] erro no fluxo de notificacao:', e.message);
    }

    const msgInterna = '🎭 *Novo pré-agendamento — Sala de Ensaio*\n\nProjeto: ' + projeto + '\nColetivo: ' + (coletivo || '-') + '\nResponsável: ' + (diretor || '-') + '\nContato: ' + (contatoNome || '-') + ' (' + whatsapp + ')\nPessoas: ' + qtdPessoas + '\n' + resumoDias + '\n\nValor: R$ ' + totalFinal.toFixed(2) + ' (sinal R$ ' + deposito.toFixed(2) + ')\nReserva: ' + reservaId;
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) { console.error('[sala-ensaio] wpp fabio:', e.message); }
    try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) { console.error('[sala-ensaio] wpp cia:', e.message); }

    res.json({ ok: true, reservaId, totalFinal, deposito });
  } catch (err) {
    console.error('[sala-ensaio] reservar-sala:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

async function atualizarStatusReserva(reservaId, novoStatus) {
  const r = await fetch('https://api.notion.com/v1/databases/' + SALA_ENSAIO_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ filter: { property: 'Reserva ID', rich_text: { equals: reservaId } }, page_size: 50 }),
  });
  const d = await r.json();
  for (const page of (d.results || [])) {
    await fetch('https://api.notion.com/v1/pages/' + page.id, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: { 'Status': { select: { name: novoStatus } } } }),
    });

    if (novoStatus === 'Cancelado') {
      const eventId = page.properties?.['Google Event ID']?.rich_text?.[0]?.plain_text || '';
      if (eventId) await excluirEventoEnsaioExterno(eventId);
    }
  }
}

// Tentativa de transferir ticket para atendimento humano (best-effort)
async function transferirParaHumano(ticketId) {
  if (!ticketId) return false;
  try {
    const r = await fetch(DIGISAC_BASE + '/tickets/' + ticketId, {
      method: 'PUT',
      headers: digisacHeaders,
      body: JSON.stringify({ departmentId: null, userId: null }),
    });
    return r.ok;
  } catch (e) {
    console.error('[sala-ensaio] erro ao transferir ticket:', e.message);
    return false;
  }
}

// Extensao do webhook: fluxo de confirmacao da Sala de Ensaio
async function processarRespostaSalaEnsaio(numero, texto, tipo, ticketId, estado) {
  const t = (texto || '').trim().toLowerCase();

  // Cliente enviou o comprovante (imagem/documento) enquanto aguardavamos
  if (estado.estado === 'aguardando_comprovante_imagem' && (tipo === 'image' || tipo === 'document')) {
    const transferido = await transferirParaHumano(ticketId);
    await enviarWhatsApp(numero, 'Recebemos seu comprovante! ✅ Nossa equipe vai verificar e confirmar seu ensaio em breve.');
    const msgInterna = '💰 *Comprovante recebido* — Reserva ' + estado.reservaId + ' (' + estado.nome + ')\nVerifique o pagamento e confirme manualmente no Notion.' + (transferido ? '\n(Ticket transferido para atendimento humano)' : '\n(Não foi possível transferir o ticket automaticamente — verifique o chat no Digisac)');
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
    try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}
    delete CONVERSAS_ESTADO[numero];
    return true;
  }

  if (estado.estado === 'aguardando_confirmacao_ensaio') {
    if (t === '1' || t.includes('confirmar')) {
      await atualizarStatusReserva(estado.reservaId, 'Aguardando Comprovante');
      await enviarWhatsApp(numero, 'Perfeito, ' + estado.nome + '! ✅\n\nAguardamos o comprovante do sinal (R$ ' + estado.deposito.toFixed(2) + ') por aqui — pode mandar a foto ou print do PIX.');
      CONVERSAS_ESTADO[numero] = { ...estado, estado: 'aguardando_comprovante_imagem' };
      const msgInterna = '✅ Cliente confirmou intenção — Reserva ' + estado.reservaId + '. Aguardando comprovante de R$ ' + estado.deposito.toFixed(2) + '.';
      try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
      try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}
      return true;
    }

    if (t === '2' || t.includes('cancelar')) {
      await atualizarStatusReserva(estado.reservaId, 'Cancelado');
      await enviarWhatsApp(numero, 'Tudo bem, ' + estado.nome + '. Seu agendamento foi cancelado. Se quiser remarcar, é só nos chamar novamente! 💛');
      const msgInterna = '❌ Cliente cancelou — Reserva ' + estado.reservaId + '.';
      try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
      try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}
      delete CONVERSAS_ESTADO[numero];
      return true;
    }

    if (t === '3' || t.includes('atendente') || t.includes('falar')) {
      const transferido = await transferirParaHumano(ticketId);
      await enviarWhatsApp(numero, 'Combinado! Alguém da nossa equipe vai te chamar por aqui em breve. 🙋');
      const msgInterna = '🙋 Cliente pediu atendimento humano — Reserva ' + estado.reservaId + '.' + (transferido ? '' : ' (transferencia automatica falhou, verifique o Digisac)');
      try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
      try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}
      delete CONVERSAS_ESTADO[numero];
      return true;
    }

    await enviarWhatsApp(numero, 'Desculpa, não entendi 🙏\n\nDigite apenas o número:\n1️⃣ Confirmar agendamento\n2️⃣ Cancelar agendamento\n3️⃣ Falar com atendente');
    return true;
  }

  if (estado.estado === 'aguardando_comprovante_imagem') {
    if (t === '2' || t.includes('cancelar')) {
      await atualizarStatusReserva(estado.reservaId, 'Cancelado');
      await enviarWhatsApp(numero, 'Tudo bem, ' + estado.nome + '. Seu agendamento foi cancelado.');
      const msgInterna = '❌ Cliente cancelou (aguardando comprovante) — Reserva ' + estado.reservaId + '.';
      try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
      try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}
      delete CONVERSAS_ESTADO[numero];
      return true;
    }
    if (t === '3' || t.includes('atendente') || t.includes('falar')) {
      const transferido = await transferirParaHumano(ticketId);
      await enviarWhatsApp(numero, 'Combinado! Alguém da nossa equipe vai te chamar por aqui em breve. 🙋');
      const msgInterna = '🙋 Cliente pediu atendimento humano (aguardando comprovante) — Reserva ' + estado.reservaId + '.';
      try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
      try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}
      delete CONVERSAS_ESTADO[numero];
      return true;
    }
    await enviarWhatsApp(numero, 'Estamos aguardando o comprovante do sinal (R$ ' + estado.deposito.toFixed(2) + '). Pode mandar a foto ou print por aqui! 📸');
    return true;
  }

  return false;
}

app.post('/enviar-comercial', async (req, res) => {
  const { numero, texto } = req.body;
  if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' });
  try {
    await enviarWhatsAppComHorarioComercial(numero, texto);
    return res.json({ ok: true });
  } catch (err) {
    return res.json({ ok: false, erro: err.message });
  }
});

// ============================================================
// SALA DE ENSAIO — Checagem de disponibilidade (Notion + Google Calendar)
// ============================================================
const CALENDARIOS_OCUPACAO = [
  'eacd87ac557ff61b7c39d7cf3c829472a75376313d0aad7a577f7832889bc2d6@group.calendar.google.com',
  'c85a56af75207511b955704c1ced645601c8f0091a61340225ac25711ba04400@group.calendar.google.com',
  '0f4566fdbafe0cc11043b6194538bb754cb90db5de165d85da709f1e362dc6b5@group.calendar.google.com',
  '89180f281678f29b2bb726f0ca19874e0ec7d105af0f402d7919236d1edea287@group.calendar.google.com',
  '5ab2609b22072bc96d8879c276ca408086e2e6a07ebd64ee74cef9cda01cfe34@group.calendar.google.com',
  'producaocialiquidificador@gmail.com',
];

async function obterOcupacaoGoogle(timeMinISO, timeMaxISO) {
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/calendar.readonly'],
    });
    const authClient = await auth.getClient();
    const calendar = google.calendar({ version: 'v3', auth: authClient });
    const resp = await calendar.freebusy.query({
      requestBody: {
        timeMin: timeMinISO,
        timeMax: timeMaxISO,
        items: CALENDARIOS_OCUPACAO.map(id => ({ id })),
      },
    });
    const intervalos = [];
    const cals = resp.data.calendars || {};
    Object.values(cals).forEach(cal => {
      (cal.busy || []).forEach(b => intervalos.push({ start: new Date(b.start), end: new Date(b.end) }));
    });
    return intervalos;
  } catch (e) {
    console.error('[sala-ensaio] erro Google Calendar:', e.message);
    return [];
  }
}

async function obterOcupacaoNotion(timeMinISO, timeMaxISO) {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + SALA_ENSAIO_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Início', date: { on_or_before: timeMaxISO } },
          { property: 'Fim', date: { on_or_after: timeMinISO } },
          { property: 'Status', select: { does_not_equal: 'Cancelado' } },
        ]},
        page_size: 100,
      }),
    });
    const d = await r.json();
    return (d.results || []).map(p => ({
      start: new Date(p.properties?.['Início']?.date?.start),
      end: new Date(p.properties?.['Fim']?.date?.start),
    })).filter(i => i.start && i.end && !isNaN(i.start) && !isNaN(i.end));
  } catch (e) {
    console.error('[sala-ensaio] erro Notion ocupacao:', e.message);
    return [];
  }
}

const BUFFER_MONTAGEM_MS = 15 * 60 * 1000;

function aplicarBufferIntervalos(intervalos) {
  return intervalos.map(iv => ({
    start: new Date(iv.start.getTime() - BUFFER_MONTAGEM_MS),
    end: new Date(iv.end.getTime() + BUFFER_MONTAGEM_MS),
  }));
}

async function verificarDisponibilidade(inicioISO, fimISO) {
  const [ocupGoogle, ocupNotion] = await Promise.all([
    obterOcupacaoGoogle(inicioISO, fimISO),
    obterOcupacaoNotion(inicioISO, fimISO),
  ]);
  const todas = aplicarBufferIntervalos([...ocupGoogle, ...ocupNotion]);
  const inicioAlvo = new Date(inicioISO);
  const fimAlvo = new Date(fimISO);
  const conflito = todas.some(iv => iv.start < fimAlvo && iv.end > inicioAlvo);
  if (conflito) return { disponivel: false, ehResidente: false };

  const dataStr = inicioISO.split('T')[0];
  const sobrepoeResidente = await verificarSobreposicaoResidente(inicioISO, fimISO, RESIDENTE_CIA_PLA_CALENDAR);
  if (sobrepoeResidente) {
    const podeUsar = await residenteDisponivelParaTerceiro(dataStr);
    if (!podeUsar) return { disponivel: false, ehResidente: true, motivoResidente: true };
    return { disponivel: true, ehResidente: true };
  }

  return { disponivel: true, ehResidente: false };
}

// GET /disponibilidade-mes?ano=2026&mes=8
app.get('/disponibilidade-mes', async (req, res) => {
  const { ano, mes } = req.query;
  if (!ano || !mes) return res.status(400).json({ error: 'ano e mes obrigatorios' });

  try {
    const anoNum = parseInt(ano, 10);
    const mesNum = parseInt(mes, 10);
    const timeMinISO = new Date(Date.UTC(anoNum, mesNum - 1, 1, 3, 0, 0)).toISOString();
    const timeMaxISO = new Date(Date.UTC(anoNum, mesNum, 1, 2, 59, 59)).toISOString();

    const [ocupGoogle, ocupNotion] = await Promise.all([
      obterOcupacaoGoogle(timeMinISO, timeMaxISO),
      obterOcupacaoNotion(timeMinISO, timeMaxISO),
    ]);
    const todasOcupacoes = aplicarBufferIntervalos([...ocupGoogle, ...ocupNotion]);

    const diasNoMes = new Date(anoNum, mesNum, 0).getDate();
    const porDia = {};
    for (let d = 1; d <= diasNoMes; d++) {
      const diaStr = ano + '-' + String(mesNum).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      porDia[diaStr] = [];
    }

    todasOcupacoes.forEach(intervalo => {
      const inicioBr = new Date(intervalo.start.getTime() - 3 * 60 * 60000);
      const fimBr = new Date(intervalo.end.getTime() - 3 * 60 * 60000);
      const diaStr = inicioBr.toISOString().split('T')[0];
      if (porDia[diaStr] !== undefined) {
        const horaIni = inicioBr.toISOString().split('T')[1].slice(0, 5);
        const horaFim = fimBr.toISOString().split('T')[1].slice(0, 5);
        porDia[diaStr].push({ inicio: horaIni, fim: horaFim });
      }
    });

    Object.keys(porDia).forEach(dia => {
      const intervalos = porDia[dia].sort((a, b) => a.inicio.localeCompare(b.inicio));
      const mesclados = [];
      intervalos.forEach(iv => {
        if (mesclados.length && iv.inicio <= mesclados[mesclados.length - 1].fim) {
          if (iv.fim > mesclados[mesclados.length - 1].fim) mesclados[mesclados.length - 1].fim = iv.fim;
        } else {
          mesclados.push({ ...iv });
        }
      });
      porDia[dia] = mesclados;
    });

    res.json({ ok: true, ano: anoNum, mes: mesNum, dias: porDia });
  } catch (err) {
    console.error('[sala-ensaio] disponibilidade-mes:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ============================================================
// SALA DE ENSAIO — Escrita no Google Calendar (Ensaio Externo)
// ============================================================
const CALENDARIO_ENSAIO_EXTERNO = '0f4566fdbafe0cc11043b6194538bb754cb90db5de165d85da709f1e362dc6b5@group.calendar.google.com';

async function getGoogleCalendarClient() {
  const { google } = require('googleapis');
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/calendar'],
  });
  const authClient = await auth.getClient();
  return google.calendar({ version: 'v3', auth: authClient });
}

async function criarEventoEnsaioExterno({ projeto, coletivo, diretor, whatsapp, tipoEnsaio, inicioISO, fimISO, reservaId }) {
  try {
    const calendar = await getGoogleCalendarClient();
    const resp = await calendar.events.insert({
      calendarId: CALENDARIO_ENSAIO_EXTERNO,
      requestBody: {
        summary: (coletivo || projeto) + ' — ' + (tipoEnsaio || 'Ensaio'),
        description: 'Responsável: ' + (diretor || '-') + '\nWhatsApp: ' + whatsapp + '\nProjeto: ' + projeto + '\nReserva: ' + reservaId,
        start: { dateTime: inicioISO },
        end: { dateTime: fimISO },
      },
    });
    return resp.data.id;
  } catch (e) {
    console.error('[sala-ensaio] erro ao criar evento no Calendar:', e.message);
    return null;
  }
}

async function excluirEventoEnsaioExterno(eventId) {
  if (!eventId) return;
  try {
    const calendar = await getGoogleCalendarClient();
    await calendar.events.delete({ calendarId: CALENDARIO_ENSAIO_EXTERNO, eventId });
  } catch (e) {
    console.error('[sala-ensaio] erro ao excluir evento do Calendar:', e.message);
  }
}


// Verifica a cada 15 minutos se algum pre-agendamento de sala ficou sem resposta por 3h+ (em horario comercial)
setInterval(async () => {
  const agora = Date.now();
  const dentroHorarioComercial = calcularProximoHorarioComercial() === null;
  if (!dentroHorarioComercial) return;

  for (const numero in CONVERSAS_ESTADO) {
    const estado = CONVERSAS_ESTADO[numero];
    if (estado.estado === 'aguardando_confirmacao_ensaio' && !estado.lembreteEnviado && estado.criadoEm) {
      if (agora - estado.criadoEm >= 3 * 60 * 60 * 1000) {
        try {
          await enviarWhatsApp(numero, 'Oi, ' + estado.nome + '! Ainda estamos aguardando sua confirmação sobre o ensaio 🎭\n\nDigite apenas o número:\n1️⃣ Confirmar agendamento\n2️⃣ Cancelar agendamento\n3️⃣ Falar com atendente');
          estado.lembreteEnviado = true;
        } catch (e) {
          console.error('[sala-ensaio] erro ao enviar lembrete de 3h:', e.message);
        }
      }
    }
  }
}, 15 * 60 * 1000);

// ============================================================
// GRUPOS RESIDENTES — Detecção dinâmica via Google Calendar
// ============================================================
const RESIDENTE_CIA_PLA_CALENDAR = 'ed190e635321c23ff3c66a1d478aa453398590a9cb97c215f3935580ba1f8480@group.calendar.google.com';
const RESIDENTES_REMARCACOES_DB = 'a743e0e7-fddc-4676-ab04-70746c01a6aa';
const WHATSAPP_CIA_PLA = '5511981578744';

async function verificarSobreposicaoResidente(inicioISO, fimISO, calendarId) {
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/calendar.readonly'],
    });
    const authClient = await auth.getClient();
    const calendar = google.calendar({ version: 'v3', auth: authClient });
    const resp = await calendar.freebusy.query({
      requestBody: { timeMin: inicioISO, timeMax: fimISO, items: [{ id: calendarId }] },
    });
    const busy = resp.data.calendars?.[calendarId]?.busy || [];
    return busy.length > 0;
  } catch (e) {
    console.error('[residentes] erro ao checar sobreposicao:', e.message);
    return false;
  }
}

// Verifica se o horario da Cia Pla pode ser tomado por um terceiro (dentro da cota) ou se deve ser BLOQUEADO (cota esgotada)
async function residenteDisponivelParaTerceiro(dataStr) {
  const semanaStr = getNumeroSemanaISO(dataStr);
  const mesStr = dataStr.slice(0, 7);
  const { totalMes, naMesmaSemana } = await contarRemarcacoesResidente('Cia Plá', mesStr, semanaStr);
  const cotaEsgotada = totalMes >= 4 || naMesmaSemana >= 1;
  return !cotaEsgotada; // true = ainda pode tomar o horario deles; false = protegido, bloquear
}

function getNumeroSemanaISO(dataStr) {
  const d = new Date(dataStr + 'T12:00:00Z');
  const target = new Date(d.valueOf());
  const dayNr = (d.getUTCDay() + 6) % 7;
  target.setUTCDate(target.getUTCDate() - dayNr + 3);
  const firstThursday = new Date(Date.UTC(target.getUTCFullYear(), 0, 4));
  const diff = target - firstThursday;
  return target.getUTCFullYear() + '-W' + (1 + Math.round(diff / (7 * 24 * 60 * 60 * 1000)));
}

async function contarRemarcacoesResidente(residente, mesStr, semanaStr) {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + RESIDENTES_REMARCACOES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Residente', select: { equals: residente } },
          { property: 'Mês', rich_text: { equals: mesStr } },
          { property: 'Status', select: { does_not_equal: 'Expirada' } },
        ]},
        page_size: 50,
      }),
    });
    const d = await r.json();
    const registros = d.results || [];
    const totalMes = registros.length;
    const naMesmaSemana = registros.filter(p => {
      const dataOriginal = p.properties?.['Data Original']?.date?.start || '';
      return dataOriginal && getNumeroSemanaISO(dataOriginal) === semanaStr;
    }).length;
    return { totalMes, naMesmaSemana };
  } catch (e) {
    console.error('[residentes] erro ao contar remarcacoes:', e.message);
    return { totalMes: 0, naMesmaSemana: 0 };
  }
}

async function registrarOfertaRemarcacao(residente, dataOriginal) {
  const mesStr = dataOriginal.slice(0, 7);
  await fetch('https://api.notion.com/v1/pages', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      parent: { database_id: RESIDENTES_REMARCACOES_DB },
      properties: {
        'Título': { title: [{ text: { content: residente + ' — oferta ' + dataOriginal } }] },
        'Residente': { select: { name: residente } },
        'Data Original': { date: { start: dataOriginal } },
        'Mês': { rich_text: [{ text: { content: mesStr } }] },
        'Status': { select: { name: 'Oferecida' } },
      },
    }),
  });
}

async function notificarResidenteSobreposicao(blocosAfetados) {
  for (const b of blocosAfetados) {
    const dataFmt = b.data.split('-').reverse().join('/');
    const semanaStr = getNumeroSemanaISO(b.data);
    const mesStr = b.data.slice(0, 7);

    const { totalMes, naMesmaSemana } = await contarRemarcacoesResidente('Cia Plá', mesStr, semanaStr);

    await registrarOfertaRemarcacao('Cia Plá', b.data);

    const limiteAtingido = totalMes >= 4 || naMesmaSemana >= 1;
    const avisoLimite = limiteAtingido
      ? '\n\n⚠️ Atenção: o limite de remarcações do mês (4x, sendo só 1x por semana) já foi atingido ou está no limite. Fale com a gente para verificar.'
      : '\n\nVocês já usaram ' + totalMes + ' de 4 remarcações este mês.';

    const msg = 'Olá! 🎭\n\nO horário de ensaio de vocês do dia ' + dataFmt + ' (' + b.inicio + ' às ' + b.fim + ') foi reservado por um cliente pagante, conforme nosso acordo.\n\nVocês podem remarcar essas horas em outro dia disponível este mês:\nhttps://agende-ensaio.ciadoliquidificador.com.br' + avisoLimite;

    try {
      await enviarWhatsAppComHorarioComercial(WHATSAPP_CIA_PLA, msg);
    } catch (e) {
      console.error('[residentes] erro ao notificar Cia Pla:', e.message);
    }
  }
}
