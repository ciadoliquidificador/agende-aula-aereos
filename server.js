const express = require('express');
const fetch = require('node-fetch');
const crypto = require('crypto');
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
const ACERVO_DB = '2b537f87-513e-45a9-b09e-28b3ff410858';
const APRESENTACOES_2026_DB = '2b9c4503-1f73-8082-8d34-f47353b066e7';

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

async function criarSubpastaDentroDe(token, parentFolderId, nomePasta) {
  const nomeSeguro = nomePasta.replace(/[<>:"/\\|?*]/g, '-').slice(0, 100);

  const createUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + parentFolderId + '/children';
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
    const getUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + parentFolderId + ':/' + encodeURIComponent(nomeSeguro);
    const getResp = await fetch(getUrl, { headers: { 'Authorization': 'Bearer ' + token } });
    if (getResp.ok) {
      const data = await getResp.json();
      return data.id;
    }
  }

  const t = await createResp.text();
  throw new Error('Criar subpasta dentro de: ' + t);
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
const APROVACOES_ALUNAS_DB = 'ec88dc5a5d9048a8b20d90332d96aa54';

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

// Envia mensagem com imagem anexada de verdade (nao so um link no texto).
// fileUrl precisa ser um link de download direto do conteudo (ex: @microsoft.graph.downloadUrl do OneDrive).
async function enviarWhatsAppComImagem(numero, texto, fileUrl) {
  const contactId = await getOrCreateContactId(numero);
  const response = await fetch(`${DIGISAC_BASE}/messages`, {
    method: 'POST',
    headers: digisacHeaders,
    body: JSON.stringify({ text: texto, fileUrl: fileUrl, type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot' })
  });
  if (!response.ok) { const t = await response.text(); throw new Error(`Digisac ${response.status}: ${t}`); }
  return response.json();
}

async function calcularProximoHorarioComercial() {
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Sao_Paulo',
    hour: 'numeric', hour12: false,
    weekday: 'short',
    year: 'numeric', month: '2-digit', day: '2-digit',
  });
  const diasMap = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };

  function partesDe(data) {
    const parts = fmt.formatToParts(data);
    const obj = {};
    parts.forEach(p => { obj[p.type] = p.value; });
    return {
      hora: parseInt(obj.hour, 10),
      diaSemana: diasMap[obj.weekday],
      dataStr: obj.year + '-' + obj.month + '-' + obj.day,
      ano: obj.year,
    };
  }

  const agora = new Date();
  const p0 = partesDe(agora);
  const feriadosCache = { [p0.ano]: await getFeriadosDoAno(p0.ano) };
  const dentroHorario = p0.hora >= 8 && p0.hora < 18 && p0.diaSemana >= 1 && p0.diaSemana <= 5 && !feriadosCache[p0.ano].has(p0.dataStr);
  if (dentroHorario) return null;

  // Avanca hora a hora ate encontrar 8h de um dia util (Brasilia), pulando feriados
  let candidato = new Date(agora);
  for (let i = 0; i < 24 * 10; i++) {
    candidato = new Date(candidato.getTime() + 60 * 60000);
    const pc = partesDe(candidato);
    if (!feriadosCache[pc.ano]) feriadosCache[pc.ano] = await getFeriadosDoAno(pc.ano);
    if (pc.hora === 8 && pc.diaSemana >= 1 && pc.diaSemana <= 5 && !feriadosCache[pc.ano].has(pc.dataStr)) {
      return candidato;
    }
  }
  return null;
}

async function enviarWhatsAppComHorarioComercial(numero, texto) {
  const agendamento = await calcularProximoHorarioComercial();
  console.log('[horario-comercial] agora=' + new Date().toISOString() + ' agendamento=' + (agendamento ? agendamento.toISOString() : 'null (envio imediato)'));
  if (agendamento) {
    await agendarMensagemFila(numero, texto, agendamento.toISOString());
    return { agendado: true };
  }
  await enviarWhatsApp(numero, texto);
  return { agendado: false };
}

const FILA_MENSAGENS_DB = '633583f0-c5b0-4e4c-81a0-48fdbd3db891';

async function agendarMensagemFila(numero, texto, enviarEmISO) {
  await fetch('https://api.notion.com/v1/pages', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      parent: { database_id: FILA_MENSAGENS_DB },
      properties: {
        'Título': { title: [{ text: { content: numero + ' — ' + enviarEmISO } }] },
        'Número': { rich_text: [{ text: { content: numero } }] },
        'Texto': { rich_text: [{ text: { content: texto } }] },
        'Enviar Em': { date: { start: enviarEmISO } },
        'Enviado': { checkbox: false },
      },
    }),
  });
}

setInterval(async () => {
  try {
    const agora = new Date().toISOString();
    const r = await fetch('https://api.notion.com/v1/databases/' + FILA_MENSAGENS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Enviado', checkbox: { equals: false } },
          { property: 'Enviar Em', date: { on_or_before: agora } },
        ]},
        page_size: 20,
      }),
    });
    const d = await r.json();
    for (const page of (d.results || [])) {
      const numero = page.properties['Número']?.rich_text?.[0]?.plain_text || '';
      const texto = page.properties['Texto']?.rich_text?.[0]?.plain_text || '';
      try {
        if (numero && texto) await enviarWhatsApp(numero, texto);
        console.log('[fila-mensagens] enviada para ' + numero);
      } catch (e) {
        console.error('[fila-mensagens] erro ao enviar para ' + numero + ':', e.message);
      }
      try {
        await fetch('https://api.notion.com/v1/pages/' + page.id, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
          body: JSON.stringify({ properties: { 'Enviado': { checkbox: true } } }),
        });
      } catch (e) {
        console.error('[fila-mensagens] erro ao marcar como enviada:', e.message);
      }
    }
  } catch (e) {
    console.error('[fila-mensagens] erro ao verificar fila:', e.message);
  }
}, 60000);

app.get('/health', (req, res) => res.json({ ok: true }));

app.post('/teste-agendamento-fila', async (req, res) => {
  const { minutos } = req.body;
  const min = parseInt(minutos, 10) || 3;
  try {
    const enviarEm = new Date(Date.now() + min * 60000).toISOString();
    await agendarMensagemFila(WHATSAPP_FABIO, '🧪 Teste da FILA PROPRIA — se voce esta lendo isso ' + min + ' minuto(s) depois de pedir o teste, a fila esta funcionando! Agendado as ' + new Date().toISOString() + ' para ' + enviarEm, enviarEm);
    res.json({ ok: true, agora: new Date().toISOString(), agendadoPara: enviarEm });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/teste-agendamento-digisac', async (req, res) => {
  const { minutos } = req.body;
  const min = parseInt(minutos, 10) || 3;
  try {
    const enviarEm = new Date(Date.now() + min * 60000).toISOString();
    const contactId = await getOrCreateContactId(WHATSAPP_FABIO);
    if (!contactId) return res.json({ ok: false, erro: 'Contato nao encontrado.' });
    const response = await fetch(DIGISAC_BASE + '/messages', {
      method: 'POST', headers: digisacHeaders,
      body: JSON.stringify({
        text: '🧪 Teste de agendamento — se voce esta lendo isso ' + min + ' minuto(s) depois de eu pedir o teste, o scheduledAt do Digisac esta funcionando! Enviado as ' + new Date().toISOString() + ', agendado para ' + enviarEm,
        type: 'chat', serviceId: SERVICE_ID, contactId, userId: USER_ID, origin: 'bot',
        scheduledAt: enviarEm,
      }),
    });
    const data = await response.json();
    res.json({ ok: response.ok, agora: new Date().toISOString(), agendadoPara: enviarEm, respostaDigisac: data });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});


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
    const { nome, telefone, turma, professor, dia, horario, data, tipo, cpf } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });

    let creditoReposicaoId = null;
    if (tipo === 'reposicao') {
      if (!cpf) return res.status(400).json({ ok: false, erro: 'CPF é obrigatório para agendar reposição.' });
      const cpfLimpo = cpf.replace(/\D/g, '');
      const cotaInfo = await verificarCotaReposicao(cpfLimpo, 'Aéreos');
      if (!cotaInfo.podeAgendar) {
        return res.status(400).json({ ok: false, erro: cotaInfo.mensagemBloqueio });
      }
      creditoReposicaoId = cotaInfo.creditos[0].id;
    }

    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: { Nome: { title: [{ text: { content: nome } }] }, Contato: { phone_number: telefone }, Turma: { select: { name: turma } }, Professor: { select: { name: professor || '' } }, Dia: { select: { name: dia || '' } }, Horário: { select: { name: horario || '' } }, Modalidade: { select: { name: 'Aéreos' } }, Status: { select: { name: 'Experimental' } }, Observações: { rich_text: [{ text: { content: `Aula experimental agendada para ${data || ''}` } }] } } }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }

    if (creditoReposicaoId) {
      await consumirCreditoReposicao(creditoReposicaoId, { turmaDesejada: turma, dataDesejada: data });
    }

    return res.json({ ok: true });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

app.post('/enviar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/agendar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/notificar', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });
app.post('/lembrete', async (req, res) => { const { numero, texto } = req.body; if (!numero || !texto) return res.json({ ok: false, erro: 'Campos obrigatorios.' }); try { await enviarWhatsApp(numero, texto); return res.json({ ok: true }); } catch (err) { return res.json({ ok: false, erro: err.message }); } });

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Proxy rodando na porta ${PORT}`));

const LIMITES_TURMAS_AEREOS = {
  "Segunda 18h": 5, "Segunda 19h": 5, "Sexta 18h": 5,
  "Terça 8h": 5, "Terça 9h": 5, "Quarta 18h": 5, "Quarta 19h": 5, "Quinta 8h": 5,
};

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

async function contarAtivasNaTurma(modalidade, turmaNome) {
  const r = await fetch(`https://api.notion.com/v1/databases/${ALUNAS_DB}/query`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { and: [
        { property: 'Status', select: { equals: 'Ativa' } },
        { property: 'Modalidade', select: { equals: modalidade } },
        { property: 'Turma', select: { equals: turmaNome } },
      ]},
      page_size: 100,
    }),
  });
  const d = await r.json();
  return (d.results || []).length;
}

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

    // Turmas que ja estao no limite de alunas ATIVAS (matriculadas) ficam bloqueadas
    // para aula experimental, independente do dia escolhido.
    const turmasLotadas = [];
    for (const [turmaNome, limite] of Object.entries(LIMITES_TURMAS_AEREOS)) {
      const ativas = await contarAtivasNaTurma('Aéreos', turmaNome);
      if (ativas >= limite) turmasLotadas.push(nomeParaId[turmaNome]);
    }

    return res.json({ ok: true, ocupadas, turmasLotadas });
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
    await agendarMensagemFila(numero, texto, new Date(enviarEm).toISOString());
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

    const ativasAcro = await contarAtivasNaTurma('Circo - Acrobacia', 'Segunda 10h');
    const turmasLotadas = ativasAcro >= 20 ? ['a1'] : [];

    return res.json({ ok: true, ocupadas, turmasLotadas });
  } catch (err) { return res.json({ ok: false, erro: err.message }); }
});

// ── Circo Infantil ────────────────────────────────────────────────────────────
app.post('/inscricao-infantil', async (req, res) => {
  try {
    const { nome, telefone, turma, professor, dia, horario, data, tipo, cpf } = req.body;
    if (!nome || !telefone || !turma) return res.json({ ok: false, erro: 'Campos obrigatórios: nome, telefone, turma.' });
    const status = tipo === 'reposicao' ? 'Ativo' : 'Experimental';

    let creditoReposicaoId = null;
    if (tipo === 'reposicao') {
      if (!cpf) return res.status(400).json({ ok: false, erro: 'CPF é obrigatório para agendar reposição.' });
      const cpfLimpo = cpf.replace(/\D/g, '');
      const cotaInfo = await verificarCotaReposicao(cpfLimpo, 'Circo Infantil');
      if (!cotaInfo.podeAgendar) {
        return res.status(400).json({ ok: false, erro: cotaInfo.mensagemBloqueio });
      }
      creditoReposicaoId = cotaInfo.creditos[0].id;
    }

    const response = await fetch(`https://api.notion.com/v1/pages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: ALUNAS_DB }, properties: { Nome: { title: [{ text: { content: nome } }] }, Contato: { phone_number: telefone }, Turma: { select: { name: turma } }, Professor: { select: { name: 'Titzi' } }, Dia: { select: { name: dia || 'Terça' } }, Horário: { select: { name: horario || '18:00' } }, Modalidade: { select: { name: 'Circo Infantil' } }, Status: { select: { name: status } }, Observações: { rich_text: [{ text: { content: `Agendado para ${data || ''}` } }] } } }),
    });
    if (!response.ok) { const t = await response.text(); throw new Error(`Notion ${response.status}: ${t}`); }

    if (creditoReposicaoId) {
      await consumirCreditoReposicao(creditoReposicaoId, { turmaDesejada: turma, dataDesejada: data });
    }

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

    const ativasTerca = await contarAtivasNaTurma('Circo Infantil', 'Terça 18h');
    const ativasQuarta = await contarAtivasNaTurma('Circo Infantil', 'Quarta 9h30');
    const turmasLotadas = [];
    if (ativasTerca >= 10) turmasLotadas.push('ci1');
    if (ativasQuarta >= 10) turmasLotadas.push('ci2');

    return res.json({ ok: true, ocupadas, turmasLotadas });
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

    // Grava também no banco Alunas principal, pra Ocupação de Turmas (Portal Admin)
    // e outras rotinas do sistema conseguirem enxergar essa matrícula normalmente.
    try {
      const valorEscolhido = (formaPagamento || '').includes('900') ? 900 : 1100;
      await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: ALUNAS_DB },
          properties: {
            'Nome': { title: [{ text: { content: nomeCompleto } }] },
            'CPF': { rich_text: [{ text: { content: cpf } }] },
            'Contato': { phone_number: telefone },
            'Email': { email: email || null },
            'Modalidade': { select: { name: 'Percussão Coletiva' } },
            'Turma': { select: { name: 'Terça 19h30' } },
            'Dia': { select: { name: 'Terça' } },
            'Horário': { select: { name: '19:30' } },
            'Professor': { select: { name: 'Lua Oliveira' } },
            'Plano': { select: { name: 'Semestral' } },
            'Frequência': { select: { name: '1x semana' } },
            'Valor': { number: valorEscolhido },
            'Status': { select: { name: 'Ativa' } },
          },
        }),
      });
    } catch (eAlunas) {
      console.error('[inscricao-percussao] erro ao gravar em Alunas (dual-write):', eAlunas.message);
    }

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

    // Dual-write no banco Alunas principal, pra Ocupação de Turmas enxergar essa matrícula.
    try {
      const valorEscolhidoCommedia = (formaPagamento || '').includes('Pix') ? 290 : 330;
      await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: ALUNAS_DB },
          properties: {
            'Nome': { title: [{ text: { content: nomeCompleto } }] },
            'CPF': { rich_text: [{ text: { content: cpf } }] },
            'Contato': { phone_number: telefone },
            'Email': { email: email || null },
            'Modalidade': { select: { name: 'Commedia Dell Arte' } },
            'Turma': { select: { name: 'Sexta 10h' } },
            'Dia': { select: { name: 'Sexta' } },
            'Horário': { select: { name: '10:00' } },
            'Professor': { select: { name: 'Diego Domingues' } },
            'Plano': { select: { name: 'Workshop' } },
            'Frequência': { select: { name: '1x semana' } },
            'Valor': { number: valorEscolhidoCommedia },
            'Status': { select: { name: 'Ativa' } },
          },
        }),
      });
    } catch (eAlunasCommedia) {
      console.error('[inscricao-commedia] erro ao gravar em Alunas (dual-write):', eAlunasCommedia.message);
    }

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
// PESSOAS — Cadastro compartilhado (reaproveita banco/token da Commedia)
// ============================================================
app.get('/pessoa/:cpf', async (req, res) => {
  const cpfLimpo = (req.params.cpf || '').replace(/\D/g, '');
  if (!cpfLimpo || cpfLimpo.length < 11) return res.json({ ok: true, encontrado: false });
  try {
    const r = await fetch(`https://api.notion.com/v1/databases/${NOTION_DB_COMMEDIA}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN_COMMEDIA}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { property: 'CPF', rich_text: { equals: cpfLimpo } },
        page_size: 1,
        sorts: [{ timestamp: 'created_time', direction: 'descending' }],
      }),
    });
    const data = await r.json();
    if (!data.results || data.results.length === 0) return res.json({ ok: true, encontrado: false });
    const p = data.results[0].properties;
    const pessoa = {
      nomeCompleto: p['Nome Completo']?.title?.[0]?.plain_text || '',
      nomeSocial: p['Nome Social']?.rich_text?.[0]?.plain_text || '',
      rg: p['RG']?.rich_text?.[0]?.plain_text || '',
      telefone: p['Telefone']?.phone_number || '',
      email: p['Email']?.email || '',
      cep: p['CEP']?.rich_text?.[0]?.plain_text || '',
      rua: p['Rua']?.rich_text?.[0]?.plain_text || '',
      numero: p['Número']?.rich_text?.[0]?.plain_text || '',
      complemento: p['Complemento']?.rich_text?.[0]?.plain_text || '',
      bairro: p['Bairro']?.rich_text?.[0]?.plain_text || '',
      cidade: p['Cidade']?.rich_text?.[0]?.plain_text || '',
      estado: p['Estado']?.rich_text?.[0]?.plain_text || '',
    };
    return res.json({ ok: true, encontrado: true, pessoa });
  } catch (err) {
    console.error('[pessoa] erro:', err.message);
    return res.json({ ok: false, encontrado: false, erro: err.message });
  }
});

// ============================================================
// DANCAS BRASILEIRAS — Inscricao (usa o mesmo banco da Commedia)
// ============================================================
app.get('/vagas-dancas', async (req, res) => {
  try {
    const r = await fetch(`https://api.notion.com/v1/databases/${NOTION_DB_COMMEDIA}/query`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN_COMMEDIA}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { and: [
        { property: 'Curso de Origem', select: { equals: 'Danças Brasileiras' } },
        { property: 'Status', select: { does_not_equal: 'Cancelado' } },
      ] } }),
    });
    const data = await r.json();
    const ocupadas = data.results ? data.results.length : 0;
    res.json({ ok: true, inscritos: ocupadas });
  } catch (err) {
    res.json({ ok: false, erro: err.message });
  }
});

// ===== CURSO DE MEDITAÇÃO — INSCRIÇÃO GRATUITA =====
app.get('/meditacao/vagas-disponiveis', async (req, res) => {
  const LIMITE_VAGAS_MEDITACAO = 10;
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Modalidade', select: { equals: 'Meditação' } },
          { property: 'Status', select: { equals: 'Ativa' } },
        ]},
        page_size: 100,
      }),
    });
    const d = await r.json();
    const inscritos = (d.results || []).length;
    res.json({ ok: true, limite: LIMITE_VAGAS_MEDITACAO, inscritos, vagasRestantes: Math.max(0, LIMITE_VAGAS_MEDITACAO - inscritos) });
  } catch (err) {
    console.error('[meditacao/vagas-disponiveis] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/meditacao/inscrever', async (req, res) => {
  const { nomeCompleto, nomeSocial, rg, cpf, telefone, email, rua, numero, complemento, bairro, cep, cidade, estado, assinatura, consentimentoUsoImagem } = req.body;
  if (!nomeCompleto || !cpf || !telefone || !email || !assinatura) return res.status(400).json({ error: 'Campos obrigatórios faltando' });
  const cpfLimpo = cpf.replace(/\D/g, '');
  if (cpfLimpo.length < 11) return res.status(400).json({ error: 'CPF inválido' });
  const numLimpo = telefone.replace(/\D/g, '');
  if (numLimpo.length < 11) return res.status(400).json({ error: 'Telefone inválido' });

  const LIMITE_VAGAS_MEDITACAO = 10;
  try {
    const rVagas = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Modalidade', select: { equals: 'Meditação' } },
          { property: 'Status', select: { equals: 'Ativa' } },
        ]},
        page_size: LIMITE_VAGAS_MEDITACAO + 1,
      }),
    });
    const dVagas = await rVagas.json();
    const inscritosAtuais = (dVagas.results || []).length;
    if (inscritosAtuais >= LIMITE_VAGAS_MEDITACAO) {
      return res.json({ ok: false, error: 'As vagas do curso de Meditação já se esgotaram. Fale com a gente pelo WhatsApp para entrar na lista de espera.' });
    }
  } catch (eVagas) {
    console.error('[meditacao/inscrever] erro ao checar vagas:', eVagas.message);
  }

  try {
    const notionResp = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN_COMMEDIA}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: NOTION_DB_COMMEDIA }, properties: {
        'Nome Completo':      { title:        [{ text: { content: nomeCompleto } }] },
        'Nome Social':        { rich_text:    [{ text: { content: nomeSocial || '' } }] },
        'CPF':                { rich_text:    [{ text: { content: cpfLimpo } }] },
        'RG':                 { rich_text:    [{ text: { content: rg || '' } }] },
        'Telefone':           { phone_number: telefone },
        'Email':              { email:        email },
        'CEP':                { rich_text:    [{ text: { content: cep || '' } }] },
        'Rua':                { rich_text:    [{ text: { content: rua || '' } }] },
        'Número':             { rich_text:    [{ text: { content: numero || '' } }] },
        'Complemento':        { rich_text:    [{ text: { content: complemento || '' } }] },
        'Bairro':             { rich_text:    [{ text: { content: bairro || '' } }] },
        'Cidade':             { rich_text:    [{ text: { content: cidade || '' } }] },
        'Estado':             { rich_text:    [{ text: { content: estado || '' } }] },
        'Forma de Pagamento': { select:       { name: 'Gratuito' } },
        'Status':             { select:       { name: 'Pendente' } },
        'Assinatura':         { rich_text:    [{ text: { content: assinatura } }] },
        'Observações':        { rich_text:    [{ text: { content: consentimentoUsoImagem ? 'Autorizou uso de imagem' : 'Não autorizou uso de imagem' } }] },
        'Curso de Origem':    { select:       { name: 'Meditação' } },
      }}),
    });
    if (!notionResp.ok) { const err = await notionResp.json(); console.error('[meditacao]', err); return res.status(500).json({ error: 'Erro Notion' }); }

    // Dual-write no banco Alunas principal, pra Ocupação de Turmas enxergar essa inscrição.
    try {
      await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: ALUNAS_DB },
          properties: {
            'Nome': { title: [{ text: { content: nomeCompleto } }] },
            'CPF': { rich_text: [{ text: { content: cpfLimpo } }] },
            'Contato': { phone_number: telefone },
            'Email': { email: email || null },
            'Modalidade': { select: { name: 'Meditação' } },
            'Turma': { select: { name: 'Turma Única 2026' } },
            'Professor': { select: { name: 'Bárbara Mazzola' } },
            'Plano': { select: { name: 'Gratuito' } },
            'Valor': { number: 0 },
            'Status': { select: { name: 'Ativa' } },
          },
        }),
      });
    } catch (eAlunasMeditacao) {
      console.error('[meditacao/inscrever] erro ao gravar em Alunas (dual-write):', eAlunasMeditacao.message);
    }

    const primeiroNome = nomeCompleto.split(' ')[0];
    const msgAdmin = `🧘 *Nova inscrição — Curso de Meditação*\n\n👤 ${nomeCompleto}\n📱 ${telefone}\n📧 ${email}`;
    const msgUser = `Olá, ${primeiroNome}! 🧘\n\nSua inscrição no curso gratuito *Corpo e Meditação*, com Bárbara Mazzola, foi recebida!\n\n8 encontros, em 3 momentos: Corpo Terreno, Corpo Sutil e Corpo Verdadeiro.\n\nEm breve entraremos em contato com as datas e horários. Qualquer dúvida é só responder aqui! ✨`;
    try { const adminId = await getOrCreateContactId('5511986899433'); if (adminId) await enviarWhatsApp('5511986899433', msgAdmin); } catch(e) {}
    try { await enviarWhatsApp('55' + numLimpo, msgUser); } catch(e) {}
    return res.json({ ok: true });
  } catch (err) { console.error('[meditacao]', err.message); return res.status(500).json({ error: 'Erro interno' }); }
});
// ===== FIM CURSO DE MEDITAÇÃO =====

app.post('/inscricao-dancas', async (req, res) => {
  const { nomeCompleto, nomeSocial, rg, cpf, telefone, email, rua, numero, complemento, bairro, cep, cidade, estado, formaPagamento, assinatura, observacoes } = req.body;
  if (!nomeCompleto || !cpf || !telefone || !email || !formaPagamento || !assinatura) return res.status(400).json({ error: 'Campos obrigatórios faltando' });
  const cpfLimpo = cpf.replace(/\D/g, '');
  if (cpfLimpo.length < 11) return res.status(400).json({ error: 'CPF inválido' });
  const numLimpo = telefone.replace(/\D/g, '');
  if (numLimpo.length < 11) return res.status(400).json({ error: 'Telefone inválido' });
  try {
    const notionResp = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN_COMMEDIA}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: NOTION_DB_COMMEDIA }, properties: {
        'Nome Completo':      { title:        [{ text: { content: nomeCompleto } }] },
        'Nome Social':        { rich_text:    [{ text: { content: nomeSocial || '' } }] },
        'CPF':                { rich_text:    [{ text: { content: cpfLimpo } }] },
        'RG':                 { rich_text:    [{ text: { content: rg || '' } }] },
        'Telefone':           { phone_number: telefone },
        'Email':              { email:        email },
        'CEP':                { rich_text:    [{ text: { content: cep || '' } }] },
        'Rua':                { rich_text:    [{ text: { content: rua || '' } }] },
        'Número':             { rich_text:    [{ text: { content: numero || '' } }] },
        'Complemento':        { rich_text:    [{ text: { content: complemento || '' } }] },
        'Bairro':             { rich_text:    [{ text: { content: bairro || '' } }] },
        'Cidade':             { rich_text:    [{ text: { content: cidade || '' } }] },
        'Estado':             { rich_text:    [{ text: { content: estado || '' } }] },
        'Forma de Pagamento': { select:       { name: formaPagamento } },
        'Status':             { select:       { name: 'Pendente' } },
        'Assinatura':         { rich_text:    [{ text: { content: assinatura } }] },
        'Observações':        { rich_text:    [{ text: { content: observacoes || '' } }] },
        'Curso de Origem':    { select:       { name: 'Danças Brasileiras' } },
      }}),
    });
    if (!notionResp.ok) { const err = await notionResp.json(); console.error('[dancas]', err); return res.status(500).json({ error: 'Erro Notion' }); }

    // Dual-write no banco Alunas principal, pra Ocupação de Turmas enxergar essa matrícula.
    try {
      const ehExperimentalDancas = formaPagamento === 'Aula Experimental Gratuita';
      const valorEscolhidoDancas = ehExperimentalDancas ? 0 : ((formaPagamento || '').includes('Pix') ? 900 : 1100);
      await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: ALUNAS_DB },
          properties: {
            'Nome': { title: [{ text: { content: nomeCompleto } }] },
            'CPF': { rich_text: [{ text: { content: cpfLimpo } }] },
            'Contato': { phone_number: telefone },
            'Email': { email: email || null },
            'Modalidade': { select: { name: 'Dancas Brasileiras' } },
            'Turma': { select: { name: 'Quarta 20h' } },
            'Dia': { select: { name: 'Quarta' } },
            'Horário': { select: { name: '20:00' } },
            'Professor': { select: { name: 'Roberta Viana' } },
            'Plano': { select: { name: ehExperimentalDancas ? 'Experimental' : 'Semestral' } },
            'Frequência': { select: { name: '1x semana' } },
            'Valor': { number: valorEscolhidoDancas },
            'Status': { select: { name: ehExperimentalDancas ? 'Experimental' : 'Ativa' } },
          },
        }),
      });
    } catch (eAlunasDancas) {
      console.error('[inscricao-dancas] erro ao gravar em Alunas (dual-write):', eAlunasDancas.message);
    }

    const primeiroNome = nomeCompleto.split(' ')[0];
    const isPix = formaPagamento.includes('Pix');
    const isExperimental = formaPagamento === 'Aula Experimental Gratuita';
    const msgAdmin = `💃 *Nova inscrição — Danças Brasileiras*\n\n👤 ${nomeCompleto}\n📱 ${telefone}\n📧 ${email}\n💳 ${formaPagamento}`;
    let instrucaoPagamento;
    if (isExperimental) {
      instrucaoPagamento = 'Você garantiu sua aula experimental gratuita, sem compromisso. Qualquer coisa é só chegar no horário! 🎉';
    } else if (isPix) {
      instrucaoPagamento = 'Envie R$ 900,00 via Pix para *fabio@cialiquidificador.com.br* e mande o comprovante aqui.';
    } else {
      instrucaoPagamento = 'Acesse o link de pagamento parcelado para concluir sua inscrição.';
    }
    const msgUser = `Olá, ${primeiroNome}! 💃\n\nSua inscrição no curso de *Danças Brasileiras* com Roberta Viana foi recebida!\n\n📅 De 05/08 a 16/12\n⏰ Quartas-feiras, 20h às 21h30\n\n${instrucaoPagamento}\n\nQualquer dúvida é só responder aqui! ✨`;
    try { const adminId = await getOrCreateContactId('5511986899433'); if (adminId) await enviarWhatsApp('5511986899433', msgAdmin); } catch(e) {}
    try { await enviarWhatsApp('55' + numLimpo, msgUser); } catch(e) {}
    return res.json({ ok: true });
  } catch (err) { console.error('[dancas]', err.message); return res.status(500).json({ error: 'Erro interno' }); }
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

// ============================================================
// APRESENTACOES — Sincronizacao com Google Calendar de verdade
// (disparado por Automation do Notion quando Data/Horario/Local mudam)
// ============================================================
const CALENDARIO_APRESENTACOES = '2c8e893b7c567c33ccbcd272b996d7e732a506ec84485e22908e36ffdf1dc999@group.calendar.google.com'; // ex: algumacoisa@group.calendar.google.com

// Funcao reaproveitavel: sincroniza UMA apresentacao (por pageId) com o Google Calendar.
// Usada pelo webhook de Apresentacoes (edicao manual), pelo webhook de Proposta aprovada
// (gatilho principal, mais confiavel) e pelo endpoint de backfill (resolver o passado 1x).
async function sincronizarApresentacaoComCalendar(pageId) {
  const rPage = await fetch('https://api.notion.com/v1/pages/' + pageId, {
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
  });
  const pageData = await rPage.json();
  const p = pageData.properties || {};

  const dataStr = p['Data da Apresentação']?.date?.start || '';
  if (!dataStr) {
    console.log('[sincronizar-apresentacao] pagina sem Data da Apresentacao, ignorando: ' + pageId);
    return null;
  }
  const dataSimples = dataStr.split('T')[0];

  const horarioTexto = p['Horário Apresentação']?.rich_text?.[0]?.plain_text || '';
  const localTitle = p['LOCAL']?.title?.[0]?.plain_text || '';
  const localPlace = p['Endereço']?.place || null;

  async function nomeTituloDaPaginaRel(id) {
    try {
      const rp = await fetch('https://api.notion.com/v1/pages/' + id, {
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
      });
      if (!rp.ok) return '';
      const pd = await rp.json();
      for (const key in (pd.properties || {})) {
        if (pd.properties[key].type === 'title') {
          return (pd.properties[key].title?.[0]?.plain_text || '').trim();
        }
      }
      return '';
    } catch (e) { return ''; }
  }

  const trabalhoRel = p['🎭 Trabalhos']?.relation || [];
  const trabalhoNome = trabalhoRel.length ? await nomeTituloDaPaginaRel(trabalhoRel[0].id) : '';

  function extrairHorarios(texto) {
    const matches = [...(texto || '').matchAll(/(\d{1,2})h(\d{2})?/g)];
    if (matches.length === 0) return null;
    const toHHMM = (m) => (m[1].padStart(2, '0')) + ':' + (m[2] || '00');
    return { inicio: toHHMM(matches[0]), fim: matches.length > 1 ? toHHMM(matches[1]) : null };
  }
  const horarios = extrairHorarios(horarioTexto);

  let eventStart, eventEnd;
  if (horarios) {
    const inicioISO = dataSimples + 'T' + horarios.inicio + ':00-03:00';
    eventStart = { dateTime: inicioISO };
    if (horarios.fim) {
      eventEnd = { dateTime: dataSimples + 'T' + horarios.fim + ':00-03:00' };
    } else {
      const fimDate = new Date(new Date(inicioISO).getTime() + 2 * 60 * 60000);
      eventEnd = { dateTime: fimDate.toISOString() };
    }
  } else {
    eventStart = { date: dataSimples };
    const proximoDia = new Date(dataSimples + 'T00:00:00Z');
    proximoDia.setUTCDate(proximoDia.getUTCDate() + 1);
    eventEnd = { date: proximoDia.toISOString().split('T')[0] };
  }

  const nomeLocal = localPlace?.name || localTitle || 'Local a definir';
  const enderecoLocal = localPlace?.address || '';
  const summary = (trabalhoNome || 'Apresentação') + ' — ' + nomeLocal;
  const description = 'Local: ' + nomeLocal + (enderecoLocal ? ' (' + enderecoLocal + ')' : '') +
    (horarioTexto ? ('\nHorário: ' + horarioTexto) : '') +
    '\nGerado automaticamente a partir do Notion.';

  const calendar = await getGoogleCalendarClient();
  const eventoExistenteId = p['Google Event ID']?.rich_text?.[0]?.plain_text || '';
  const requestBody = { summary, description, start: eventStart, end: eventEnd };
  if (enderecoLocal) requestBody.location = enderecoLocal;

  let googleEventId = eventoExistenteId;
  if (eventoExistenteId) {
    try {
      await calendar.events.update({ calendarId: CALENDARIO_APRESENTACOES, eventId: eventoExistenteId, requestBody });
      console.log('[sincronizar-apresentacao] evento atualizado: ' + eventoExistenteId);
    } catch (e) {
      console.error('[sincronizar-apresentacao] erro ao atualizar, criando novo:', e.message);
      const resp = await calendar.events.insert({ calendarId: CALENDARIO_APRESENTACOES, requestBody });
      googleEventId = resp.data.id;
    }
  } else {
    const resp = await calendar.events.insert({ calendarId: CALENDARIO_APRESENTACOES, requestBody });
    googleEventId = resp.data.id;
  }

  if (googleEventId !== eventoExistenteId) {
    await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: { 'Google Event ID': { rich_text: [{ text: { content: googleEventId } }] } } }),
    });
  }

  console.log('[sincronizar-apresentacao] concluido, evento: ' + googleEventId + ' (pagina ' + pageId + ')');
  return googleEventId;
}

app.post('/webhook-apresentacao-notion', async (req, res) => {
  res.status(200).json({ ok: true }); // responde rapido, processa depois

  try {
    const body = req.body || {};
    console.log('[webhook-apresentacao-notion] payload recebido:', JSON.stringify(body).slice(0, 500));

    const pageId = (body.data && body.data.id) || body.pageId || body.page_id || null;
    if (!pageId) {
      console.error('[webhook-apresentacao-notion] payload sem page id reconhecivel.');
      return;
    }

    await sincronizarApresentacaoComCalendar(pageId);
  } catch (err) {
    console.error('[webhook-apresentacao-notion] erro:', err.message);
  }
});

// ============================================================
// BACKFILL MANUAL (rodar 1x, sem polling continuo) — resolve o
// passado: apresentacoes que ja chegaram preenchidas e nunca
// dispararam nenhum gatilho. Chamar via curl/navegador uma vez.
// ============================================================
app.get('/admin/sincronizar-apresentacoes-pendentes', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + APRESENTACOES_2026_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Data da Apresentação', date: { is_not_empty: true } },
          { property: 'Google Event ID', rich_text: { is_empty: true } },
        ]},
        page_size: 50,
      }),
    });
    const d = await r.json();
    const pendentes = d.results || [];
    const resultado = [];
    for (const pagina of pendentes) {
      try {
        const eventId = await sincronizarApresentacaoComCalendar(pagina.id);
        resultado.push({ pageId: pagina.id, ok: true, eventId });
      } catch (e) {
        resultado.push({ pageId: pagina.id, ok: false, erro: e.message });
      }
    }
    res.json({ ok: true, totalProcessadas: pendentes.length, resultado });
  } catch (err) {
    console.error('[sincronizar-apresentacoes-pendentes] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// APRESENTACOES — Notificacao de escalacao (Producao/Elenco/Tecnicos)
// Dispara quando ELENCO, Producao Liqui, Tecnico de Som ou Tecnico de Luz sao definidos
// ============================================================
async function buscarDadosApresentacao(pageId) {
  const rPage = await fetch('https://api.notion.com/v1/pages/' + pageId, {
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
  });
  const pageData = await rPage.json();
  const p = pageData.properties || {};

  const dataStr = p['Data da Apresentação']?.date?.start || '';
  const dataSimples = dataStr ? dataStr.split('T')[0] : '';
  const dataFmt = dataSimples ? dataSimples.split('-').reverse().join('/') : '';
  const horarioTexto = p['Horário Apresentação']?.rich_text?.[0]?.plain_text || '';
  const localTitle = p['LOCAL']?.title?.[0]?.plain_text || '';
  const localPlace = p['Local']?.place || null;
  const nomeLocal = localPlace?.name || localTitle || 'Local a definir';
  const enderecoLocal = localPlace?.address || '';
  const localSaida = p['Local Saída']?.select?.name || '';
  const horarioSaida = p['Horário de Saída']?.rich_text?.[0]?.plain_text || '';

  let trabalhoNome = '';
  const trabalhoRel = p['🎭 Trabalhos']?.relation || [];
  if (trabalhoRel.length) {
    try {
      const rt = await fetch('https://api.notion.com/v1/pages/' + trabalhoRel[0].id, {
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
      });
      const td = await rt.json();
      for (const key in (td.properties || {})) {
        if (td.properties[key].type === 'title') { trabalhoNome = (td.properties[key].title?.[0]?.plain_text || '').trim(); break; }
      }
    } catch (e) {}
  }

  const idsEnvolvidos = new Set();
  ['Produção Liqui', 'ELENCO', 'TÉCNICO DE SOM', 'TÉCNICO DE LUZ'].forEach(campo => {
    (p[campo]?.relation || []).forEach(rel => idsEnvolvidos.add(rel.id));
  });

  return { p, dataFmt, horarioTexto, nomeLocal, enderecoLocal, localSaida, horarioSaida, trabalhoNome, idsEnvolvidos };
}

async function notificarPessoasApresentacao(idsEnvolvidos, montarMensagem) {
  for (const idPessoa of idsEnvolvidos) {
    try {
      const rPessoa = await fetch('https://api.notion.com/v1/pages/' + idPessoa, {
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
      });
      const pessoaData = await rPessoa.json();
      const pp = pessoaData.properties || {};
      const nomePessoa = pp['Nome']?.title?.[0]?.plain_text || '';
      const telefonePessoa = pp['Telefone']?.phone_number || '';
      if (!telefonePessoa) continue;
      const numLimpo = telefonePessoa.replace(/\D/g, '');
      const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
      const primeiroNome = (nomePessoa || '').split(' ')[0];
      await enviarWhatsAppComHorarioComercial(numBr, montarMensagem(primeiroNome));
    } catch (e) {
      console.error('[apresentacoes] erro ao notificar pessoa ' + idPessoa + ':', e.message);
    }
  }
}

app.post('/webhook-apresentacao-escalacao', async (req, res) => {
  res.status(200).json({ ok: true });
  try {
    const body = req.body || {};
    const pageId = (body.data && body.data.id) || body.pageId || body.page_id || null;
    if (!pageId) { console.error('[webhook-apresentacao-escalacao] sem page id.'); return; }

    const { p, dataFmt, horarioTexto, nomeLocal, enderecoLocal, trabalhoNome, idsEnvolvidos } = await buscarDadosApresentacao(pageId);
    if (!dataFmt) { console.log('[webhook-apresentacao-escalacao] sem data, ignorando.'); return; }

    const jaNotificadosStr = p['Escalação Notificados']?.rich_text?.[0]?.plain_text || '';
    const jaNotificados = new Set(jaNotificadosStr.split(',').map(s => s.trim()).filter(Boolean));
    const novosIds = [...idsEnvolvidos].filter(id => !jaNotificados.has(id));

    if (novosIds.length === 0) {
      console.log('[webhook-apresentacao-escalacao] nenhuma pessoa nova para notificar.');
      return;
    }

    await notificarPessoasApresentacao(novosIds, (primeiroNome) =>
      'Olá, ' + primeiroNome + '! 🎭\n\nVocê está escalado(a) para a apresentação:\n\n🎬 ' + (trabalhoNome || 'Apresentação') +
      '\n📅 ' + dataFmt + (horarioTexto ? ('\n⏰ ' + horarioTexto) : '') +
      '\n📍 ' + nomeLocal + (enderecoLocal ? (' - ' + enderecoLocal) : '') +
      '\n\nQualquer dúvida, é só chamar!'
    );

    const todosNotificados = [...jaNotificados, ...novosIds];
    await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: { 'Escalação Notificados': { rich_text: [{ text: { content: todosNotificados.join(',') } }] } } }),
    });

    console.log('[webhook-apresentacao-escalacao] concluido para pagina ' + pageId + ', novos notificados: ' + novosIds.length);
  } catch (err) {
    console.error('[webhook-apresentacao-escalacao] erro:', err.message);
  }
});

// ============================================================
// APRESENTACOES — Notificacao de saida
// Dispara quando Local Saida ou Horario de Saida sao definidos
// ============================================================
app.post('/webhook-apresentacao-saida', async (req, res) => {
  res.status(200).json({ ok: true });
  try {
    const body = req.body || {};
    const pageId = (body.data && body.data.id) || body.pageId || body.page_id || null;
    if (!pageId) { console.error('[webhook-apresentacao-saida] sem page id.'); return; }

    const { p, dataFmt, trabalhoNome, localSaida, horarioSaida, idsEnvolvidos } = await buscarDadosApresentacao(pageId);
    if (!localSaida && !horarioSaida) { console.log('[webhook-apresentacao-saida] sem local/horario de saida, ignorando.'); return; }

    const jaNotificadosStr = p['Saída Notificados']?.rich_text?.[0]?.plain_text || '';
    const jaNotificados = new Set(jaNotificadosStr.split(',').map(s => s.trim()).filter(Boolean));
    const novosIds = [...idsEnvolvidos].filter(id => !jaNotificados.has(id));

    if (novosIds.length === 0) {
      console.log('[webhook-apresentacao-saida] nenhuma pessoa nova para notificar.');
      return;
    }

    await notificarPessoasApresentacao(novosIds, (primeiroNome) =>
      'Olá, ' + primeiroNome + '! 🚐\n\nSaída definida para a apresentação ' + (trabalhoNome || '') + (dataFmt ? (' (' + dataFmt + ')') : '') + ':\n\n' +
      (localSaida ? ('📍 Local de saída: ' + localSaida + '\n') : '') +
      (horarioSaida ? ('⏰ Horário de saída: ' + horarioSaida + '\n') : '') +
      '\nNos vemos lá!'
    );

    const todosNotificados = [...jaNotificados, ...novosIds];
    await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: { 'Saída Notificados': { rich_text: [{ text: { content: todosNotificados.join(',') } }] } } }),
    });

    console.log('[webhook-apresentacao-saida] concluido para pagina ' + pageId + ', novos notificados: ' + novosIds.length);
  } catch (err) {
    console.error('[webhook-apresentacao-saida] erro:', err.message);
  }
});

// ============================================================
// PROFESSOR SUB — Fluxo de substituicao de professor
// ============================================================
const SUBSTITUICOES_DB = 'ee9e1fb2089d4734926bb8a941a08b5e';

const PROFESSORES_SUB = {
  'Aéreos': [
    { nome: 'Gabi', telefone: '5511961416621', turmas: ['Segunda 18h', 'Segunda 19h', 'Sexta 18h'] },
    { nome: 'Talita', telefone: '5511989142791', turmas: ['Terça 8h', 'Terça 9h'] },
    { nome: 'Gustra', telefone: '5511988485740', turmas: ['Quarta 18h', 'Quarta 19h'] },
    { nome: 'Guilherme', telefone: '5511989538880', turmas: ['Quinta 8h'] },
  ],
  'Acrobacia': [
    { nome: 'André', telefone: '5511981578744', turmas: ['Segunda 10h'] },
  ],
  'Circo Infantil': [
    { nome: 'Titzi', telefone: '5511951780877', turmas: ['Terça 18h', 'Quarta 9h30'] },
  ],
  'Yoga': [
    { nome: 'Giulia', telefone: '5512988222584', turmas: ['Quarta 7h', 'Quarta 8h', 'Sexta 7h', 'Sexta 8h'] },
  ],
  'Danças Brasileiras': [
    { nome: 'Roberta', telefone: '5511971918173', turmas: ['Quarta 20h'] },
  ],
};

const SUBSTITUICOES_BROADCAST = {}; // broadcastId -> { professorFaltante, modalidade, turma, data, notionPageId, telefonesConsultados: [], recusas: [], resolvido: false }

function horarioDaTurma(turmaNome) {
  const m = (turmaNome || '').match(/(\d{1,2})h(\d{2})?/);
  if (!m) return null;
  return m[1].padStart(2, '0') + ':' + (m[2] || '00');
}

async function agendarLembreteSub(numeroSubstituto, nomeSubstituto, professorFaltante, turma, modalidade, data) {
  try {
    const horario = horarioDaTurma(turma);
    if (!horario) return;
    const aulaDate = new Date(data + 'T' + horario + ':00-03:00');
    const horasAteAula = (aulaDate.getTime() - Date.now()) / (1000 * 60 * 60);
    if (horasAteAula <= 24) return; // aula e hoje/amanha mesmo, sem tempo pro lembrete de 24h

    const enviarLembreteEm = new Date(aulaDate.getTime() - 24 * 60 * 60 * 1000).toISOString();
    const primeiroNome = (nomeSubstituto || '').split(' ')[0];
    const dataFmt = data.split('-').reverse().join('/');
    const msg = 'Olá, ' + primeiroNome + '! 🔔\n\nLembrete: amanhã você vai substituir ' + professorFaltante + ' na turma de ' + turma + ' (' + modalidade + ')!\n\n📅 ' + dataFmt + '\n🏠 Local: Espaço Liquidificador\n💰 Valor: R$ 80,00 por aula, pago pelo ' + professorFaltante + ' direto pra você.\n\nQualquer dúvida, chama a gente!';

    // Usa a fila de mensagens (Notion + processamento programado) em vez do scheduledAt
    // direto do Digisac, que se mostrou nao confiavel pra esse tipo de agendamento
    // (disparava na hora em vez de esperar o horario certo).
    await agendarMensagemFila(numeroSubstituto, msg, enviarLembreteEm);
  } catch (e) {
    console.error('[sub] erro ao agendar lembrete 24h:', e.message);
  }
}

app.get('/debug-horario-comercial', async (req, res) => {
  try {
    const agora = new Date();
    const agendamento = await calcularProximoHorarioComercial();
    res.json({
      ok: true,
      agora: agora.toISOString(),
      agoraBrasilia: new Date(agora.getTime() - 3 * 60 * 60000).toISOString(),
      dentroDoHorarioComercial: agendamento === null,
      proximoEnvioAgendadoPara: agendamento ? agendamento.toISOString() : null,
    });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Pool de possiveis substitutos por modalidade, quando difere de quem da aula normalmente
// ali (ex: Circo Infantil so tem a Titzi como titular, mas qualquer professor de circo pode cobrir)
const POOL_SUBSTITUTOS_SUB = {
  'Circo Infantil': [
    { nome: 'Gustra', telefone: '5511988485740' },
    { nome: 'Guilherme', telefone: '5511989538880' },
    { nome: 'André', telefone: '5511981578744' },
    { nome: 'Talita', telefone: '5511989142791' },
    { nome: 'Gabi', telefone: '5511961416621' },
  ],
};

app.get('/professores-sub', (req, res) => {
  res.json({ ok: true, professores: PROFESSORES_SUB, poolSubstitutos: POOL_SUBSTITUTOS_SUB });
});

async function criarRegistroSubstituicao({ professorFaltante, modalidade, turma, data, status, substituto, whatsappSubstituto }) {
  const resp = await fetch('https://api.notion.com/v1/pages', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      parent: { database_id: SUBSTITUICOES_DB },
      properties: {
        'Título': { title: [{ text: { content: professorFaltante + ' — ' + turma + ' (' + data + ')' } }] },
        'Professor Titular': { rich_text: [{ text: { content: professorFaltante } }] },
        'Modalidade': { select: { name: modalidade } },
        'Turma': { rich_text: [{ text: { content: turma } }] },
        'Data da Falta': { date: { start: data } },
        'Status': { select: { name: status } },
        'Substituto': { rich_text: [{ text: { content: substituto || '' } }] },
        'WhatsApp Substituto': { phone_number: whatsappSubstituto || null },
      },
    }),
  });
  const d = await resp.json();
  return d.id;
}

async function atualizarRegistroSubstituicao(pageId, { status, substituto, whatsappSubstituto }) {
  const properties = {};
  if (status) properties['Status'] = { select: { name: status } };
  if (substituto !== undefined) properties['Substituto'] = { rich_text: [{ text: { content: substituto || '' } }] };
  if (whatsappSubstituto !== undefined) properties['WhatsApp Substituto'] = { phone_number: whatsappSubstituto || null };
  await fetch('https://api.notion.com/v1/pages/' + pageId, {
    method: 'PATCH',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ properties }),
  });
}

// Professor ja resolveu sozinho (substituto proprio) ou vai falar direto com o Fabio
app.post('/sub-resolvido', async (req, res) => {
  const { professorFaltante, modalidade, turma, data, tipo, substituto, whatsappSubstituto } = req.body;
  if (!professorFaltante || !modalidade || !turma || !data || !tipo) {
    return res.status(400).json({ ok: false, erro: 'Campos obrigatorios faltando.' });
  }
  try {
    const status = tipo === 'substituto_proprio' ? 'Resolvido - Substituto Próprio' : 'Sem Substituto - Fábio Notificado';
    const pageId = await criarRegistroSubstituicao({
      professorFaltante, modalidade, turma, data, status,
      substituto: tipo === 'substituto_proprio' ? substituto : '',
      whatsappSubstituto: tipo === 'substituto_proprio' ? whatsappSubstituto : '',
    });

    if (tipo === 'substituto_proprio' && whatsappSubstituto) {
      const numLimpoSub = whatsappSubstituto.replace(/\D/g, '');
      const numBrSub = numLimpoSub.length === 11 ? '55' + numLimpoSub : numLimpoSub;
      agendarLembreteSub(numBrSub, substituto, professorFaltante, turma, modalidade, data).catch(()=>{});

      // Confirmação imediata pro substituto, já com o valor
      const dataFmtSub = data.split('-').reverse().join('/');
      const primeiroNomeSub = (substituto || '').split(' ')[0];
      const msgConfirmaSub = 'Olá, ' + primeiroNomeSub + '! ✅\n\nVocê foi combinado(a) como substituto(a) de ' + professorFaltante + ' na turma de ' + turma + ' (' + modalidade + ') no dia ' + dataFmtSub + '.\n\n💰 Valor: R$ 80,00 por aula, pago pelo(a) ' + professorFaltante + ' direto pra você.';
      try { await enviarWhatsApp(numBrSub, msgConfirmaSub); } catch(e) {}

      // Lembrete pro titular de que ele deve pagar o substituto
      try {
        const telTitular = await buscarTelefoneProfessorPorNome(professorFaltante);
        if (telTitular) {
          const msgTitular = 'Combinado! ✅\n\nSua substituição com ' + substituto + ' na turma de ' + turma + ' (' + modalidade + ') no dia ' + dataFmtSub + ' foi registrada.\n\n💰 Lembrete: o valor de R$ 80,00 por aula deve ser pago por você diretamente ao(à) ' + substituto + '.';
          await enviarWhatsApp(telTitular, msgTitular);
        }
      } catch (e) {}
    }

    const dataFmt = data.split('-').reverse().join('/');
    let msgFabio;
    if (tipo === 'substituto_proprio') {
      msgFabio = '🔄 *Substituição resolvida*\n\nProfessor: ' + professorFaltante + '\nTurma: ' + turma + ' (' + modalidade + ')\nData: ' + dataFmt + '\nSubstituto: ' + substituto + (whatsappSubstituto ? ' (' + whatsappSubstituto + ')' : '');
    } else {
      msgFabio = '⚠️ *Preciso de ajuda com substituição*\n\nProfessor: ' + professorFaltante + '\nTurma: ' + turma + ' (' + modalidade + ')\nData: ' + dataFmt + '\n\nJá consultou os professores do espaço, ninguém pôde cobrir.';
    }
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgFabio); } catch(e) {}

    res.json({ ok: true, pageId });
  } catch (err) {
    console.error('[sub-resolvido] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// Dispara pergunta pra todos os outros professores da modalidade
app.post('/sub-broadcast', async (req, res) => {
  const { professorFaltante, modalidade, turma, data } = req.body;
  if (!professorFaltante || !modalidade || !turma || !data) {
    return res.status(400).json({ ok: false, erro: 'Campos obrigatorios faltando.' });
  }
  try {
    const outros = POOL_SUBSTITUTOS_SUB[modalidade]
      ? POOL_SUBSTITUTOS_SUB[modalidade].filter(p => p.nome !== professorFaltante)
      : (PROFESSORES_SUB[modalidade] || []).filter(p => p.nome !== professorFaltante);
    if (outros.length === 0) {
      return res.json({ ok: false, erro: 'Não há outros professores cadastrados nessa modalidade.' });
    }

    const pageId = await criarRegistroSubstituicao({
      professorFaltante, modalidade, turma, data, status: 'Aguardando Confirmação',
    });

    const broadcastId = 'SUB-' + Date.now();
    const dataFmt = data.split('-').reverse().join('/');
    const msg = 'Olá! 🎪\n\n' + professorFaltante + ' vai faltar na turma de ' + turma + ' (' + modalidade + ') no dia ' + dataFmt + '.\n\nVocê pode cobrir essa aula?\n\n💰 Valor: R$ 80,00 por aula, pago pelo(a) ' + professorFaltante + ' direto pra você.\n\nResponda *SIM* ou *NÃO*.';

    SUBSTITUICOES_BROADCAST[broadcastId] = {
      professorFaltante, modalidade, turma, data, dataFmt, notionPageId: pageId,
      telefonesConsultados: outros.map(p => p.telefone), recusas: [], resolvido: false,
    };

    for (const prof of outros) {
      CONVERSAS_ESTADO[prof.telefone] = { estado: 'aguardando_resposta_sub', broadcastId };
      try { await enviarWhatsAppComHorarioComercial(prof.telefone, msg); } catch(e) {}
    }

    res.json({ ok: true, pageId, broadcastId });
  } catch (err) {
    console.error('[sub-broadcast] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// RESIDENCIA ARTISTICA 2026 — Inscricao do chamamento
// ============================================================
const RESIDENCIA_DB = 'a514d3c11a8340ffbc73fa30582f6296';

app.post('/residencia/inscrever', async (req, res) => {
  const {
    nomeCivil, nomeSocial, email, telefone, rg, cpf, funcaoProjeto,
    nomeColetivo, qtdIntegrantes, fichaTecnica,
    nomeProjeto, linguagemArtistica, descricaoProjeto, justificativa,
    datasPretendidas, previsaoApresentacao, contempladoLei, qualLei,
    generoMulherCis, generoHomemCis, generoMulherTrans, generoHomemTrans, generoNaoBinaria, generoOutroQtd, generoOutroTexto, generoNaoDeclarar,
    etniaNegras, etniaPardas, etniaAmarelas, etniaIndigenas, etniaBrancas, etniaOutras, etniaNaoDeclarar,
    pcdQuantidade, pcdNaoDeclarar,
    curriculoProponente, curriculoResponsavel, curriculoIntegrantes, propostaContrapartida,
    materialArquivo, materialNomeArquivo,
    termoAceito,
  } = req.body;

  if (!nomeCivil || !email || !telefone || !rg || !cpf || !funcaoProjeto || !nomeColetivo ||
      !nomeProjeto || !linguagemArtistica || !descricaoProjeto || !justificativa || !datasPretendidas ||
      !curriculoProponente || !curriculoResponsavel || !propostaContrapartida || !termoAceito) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }

  try {
    let linkMaterial = '';
    if (materialArquivo) {
      const msToken = await getMicrosoftToken();
      const nomePasta = 'Residencia2026-' + slugify(nomeProjeto) + '-' + slugify(nomeColetivo);
      const folderId = await criarOuObterSubpasta(msToken, nomePasta);
      const base64Clean = materialArquivo.replace(/^data:.*;base64,/, '');
      const buffer = Buffer.from(base64Clean, 'base64');
      const nomeArquivo = materialNomeArquivo || 'material-complementar';
      const uploadUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + folderId + ':/' + encodeURIComponent(nomeArquivo) + ':/content';
      const rUpload = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { 'Authorization': 'Bearer ' + msToken, 'Content-Type': 'application/octet-stream' },
        body: buffer,
      });
      if (!rUpload.ok) { const t = await rUpload.text(); throw new Error('Upload OneDrive: ' + t); }
      const uploaded = await rUpload.json();
      linkMaterial = await criarLinkCompartilhamento(msToken, uploaded.id);
    }

    const numLimpo = (telefone || '').replace(/\D/g, '');
    const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;

    const rNotion = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: RESIDENCIA_DB },
        properties: {
          'Título': { title: [{ text: { content: nomeProjeto + ' — ' + nomeColetivo } }] },
          'Nome Civil': { rich_text: [{ text: { content: nomeCivil } }] },
          'Nome Social': { rich_text: [{ text: { content: nomeSocial || '' } }] },
          'Email': { email: email },
          'Telefone': { phone_number: telefone },
          'RG': { rich_text: [{ text: { content: rg } }] },
          'CPF': { rich_text: [{ text: { content: (cpf || '').replace(/\D/g, '') } }] },
          'Função no Projeto': { rich_text: [{ text: { content: funcaoProjeto } }] },
          'Nome do Coletivo': { rich_text: [{ text: { content: nomeColetivo } }] },
          'Quantidade de Integrantes': { number: qtdIntegrantes ? Number(qtdIntegrantes) : null },
          'Ficha Técnica': { rich_text: [{ text: { content: fichaTecnica || '' } }] },
          'Nome do Projeto': { rich_text: [{ text: { content: nomeProjeto } }] },
          'Linguagem Artística': { select: { name: linguagemArtistica } },
          'Descrição do Projeto': { rich_text: [{ text: { content: descricaoProjeto } }] },
          'Justificativa / Carta de Intenção': { rich_text: [{ text: { content: justificativa } }] },
          'Datas Pretendidas': { select: { name: datasPretendidas } },
          'Previsão de Apresentação': { rich_text: [{ text: { content: previsaoApresentacao || '' } }] },
          'Contemplado por Lei de Incentivo': { select: { name: contempladoLei || 'Não' } },
          'Qual Lei de Incentivo': { rich_text: [{ text: { content: qualLei || '' } }] },
          'Gênero - Mulher Cis': { number: generoNaoDeclarar ? null : Number(generoMulherCis || 0) },
          'Gênero - Homem Cis': { number: generoNaoDeclarar ? null : Number(generoHomemCis || 0) },
          'Gênero - Mulher Trans': { number: generoNaoDeclarar ? null : Number(generoMulherTrans || 0) },
          'Gênero - Homem Trans': { number: generoNaoDeclarar ? null : Number(generoHomemTrans || 0) },
          'Gênero - Não Binária': { number: generoNaoDeclarar ? null : Number(generoNaoBinaria || 0) },
          'Gênero - Outro (Quantidade)': { number: generoNaoDeclarar ? null : Number(generoOutroQtd || 0) },
          'Gênero - Outro (Qual)': { rich_text: [{ text: { content: generoOutroTexto || '' } }] },
          'Gênero - Prefiro Não Declarar': { checkbox: !!generoNaoDeclarar },
          'Etnia - Negras': { number: etniaNaoDeclarar ? null : Number(etniaNegras || 0) },
          'Etnia - Pardas': { number: etniaNaoDeclarar ? null : Number(etniaPardas || 0) },
          'Etnia - Amarelas': { number: etniaNaoDeclarar ? null : Number(etniaAmarelas || 0) },
          'Etnia - Indígenas': { number: etniaNaoDeclarar ? null : Number(etniaIndigenas || 0) },
          'Etnia - Brancas': { number: etniaNaoDeclarar ? null : Number(etniaBrancas || 0) },
          'Etnia - Outras': { number: etniaNaoDeclarar ? null : Number(etniaOutras || 0) },
          'Etnia - Prefiro Não Declarar': { checkbox: !!etniaNaoDeclarar },
          'PCD - Quantidade': { number: pcdNaoDeclarar ? null : Number(pcdQuantidade || 0) },
          'PCD - Prefiro Não Declarar': { checkbox: !!pcdNaoDeclarar },
          'Currículo do Proponente': { rich_text: [{ text: { content: curriculoProponente } }] },
          'Currículo do Responsável': { rich_text: [{ text: { content: curriculoResponsavel } }] },
          'Currículo dos Integrantes': { rich_text: [{ text: { content: curriculoIntegrantes || '' } }] },
          'Proposta de Contrapartida': { rich_text: [{ text: { content: propostaContrapartida } }] },
          'Link Material Complementar': { url: linkMaterial || null },
          'Termo Aceito': { checkbox: !!termoAceito },
          'Status': { select: { name: 'Recebida' } },
        },
      }),
    });

    if (!rNotion.ok) {
      const corpoErro = await rNotion.text();
      console.error('[residencia] ERRO ao gravar no Notion:', rNotion.status, corpoErro);
      throw new Error('Falha ao gravar a inscrição no Notion: ' + corpoErro);
    }

    const primeiroNome = nomeCivil.split(' ')[0];
    if (numBr) {
      try {
        await enviarWhatsApp(numBr, 'Olá, ' + primeiroNome + '! 🎪\n\nSua inscrição no *Chamamento de Residência Artística — Espaço Liquidificador (2º semestre 2026)* foi recebida com sucesso!\n\n📅 Projeto: ' + nomeProjeto + '\n👥 Coletivo: ' + nomeColetivo + '\n\nO resultado será divulgado em 08/08/2026 no Instagram @espacoliqui. Qualquer dúvida, escreva para residencia@cialiquidificador.com.br.');
      } catch(e) {}
    }
    const msgInterna = '🎪 *Nova inscrição — Residência Artística 2026*\n\nProjeto: ' + nomeProjeto + '\nColetivo: ' + nomeColetivo + '\nResponsável: ' + nomeCivil + ' (' + telefone + ')\nDatas pretendidas: ' + datasPretendidas;
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}

    res.json({ ok: true });
  } catch (err) {
    console.error('[residencia] erro ao inscrever:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// MATRICULA — Inscricao para cursos regulares (Yoga, Aereos, Acro, Infantil)
// ============================================================
const MODALIDADES_MATRICULA = {
  'Yoga': {
    turmas: [
      { nome: 'Quarta 7h', dia: 'Quarta', horario: '07:00', professor: 'Giulia', limite: 10 },
      { nome: 'Quarta 8h', dia: 'Quarta', horario: '08:00', professor: 'Giulia', limite: 10 },
      { nome: 'Sexta 7h', dia: 'Sexta', horario: '07:00', professor: 'Giulia', limite: 10 },
      { nome: 'Sexta 8h', dia: 'Sexta', horario: '08:00', professor: 'Giulia', limite: 10 },
    ],
    permiteFrequenciaDupla: true,
    precos: {
      '1x semana': { Mensal: 180.00, Semestral: 162.00, Anual: 144.00 },
      '2x semana': { Mensal: 285.00, Semestral: 256.50, Anual: 228.00 },
    },
  },
  'Aéreos': {
    turmas: [
      { nome: 'Segunda 18h', dia: 'Segunda', horario: '18:00', professor: 'Gabi', limite: 5 },
      { nome: 'Segunda 19h', dia: 'Segunda', horario: '19:00', professor: 'Gabi', limite: 5 },
      { nome: 'Terça 8h', dia: 'Terça', horario: '08:00', professor: 'Talita', limite: 5 },
      { nome: 'Terça 9h', dia: 'Terça', horario: '09:00', professor: 'Talita', limite: 5 },
      { nome: 'Quarta 18h', dia: 'Quarta', horario: '18:00', professor: 'Gustra', limite: 5 },
      { nome: 'Quarta 19h', dia: 'Quarta', horario: '19:00', professor: 'Gustra', limite: 5 },
      { nome: 'Quinta 8h', dia: 'Quinta', horario: '08:00', professor: 'Guilherme', limite: 5 },
    ],
    permiteFrequenciaDupla: true,
    precos: {
      '1x semana': { Mensal: 255.00, Semestral: 230.00, Anual: 207.00 },
      '2x semana': { Mensal: 395.00, Semestral: 355.00, Anual: 315.00 },
    },
  },
  'Circo - Acrobacia': {
    turmas: [
      { nome: 'Segunda 10h', dia: 'Segunda', horario: '10:00', professor: 'André', limite: 20 },
    ],
    permiteFrequenciaDupla: false,
    precos: {
      '1x semana': { Mensal: 200.00 },
    },
  },
  'Circo Infantil': {
    turmas: [
      { nome: 'Terça 18h', dia: 'Terça', horario: '18:00', professor: 'Titzi', limite: 10 },
      { nome: 'Quarta 9h30', dia: 'Quarta', horario: '09:30', professor: 'Titzi', limite: 10 },
    ],
    permiteFrequenciaDupla: false,
    precos: {
      '1x semana': { Mensal: 215.00, Semestral: 195.00, Anual: 175.00 },
    },
  },
};

const OBJETO_POR_MODALIDADE = {
  'Yoga': 'A CONTRATADA ministra aulas regulares de yoga, em grupo, com uma hora de duração, abrangendo posturas (asanas), técnicas de respiração (pranayama) e relaxamento, adaptadas ao nível do(a) praticante.',
  'Aéreos': 'A CONTRATADA ministra aulas regulares de circo aéreo, oferecendo aulas semanais em grupo com uma hora de duração. Em todos os planos, as(os) alunas(os) utilizam todos os aparelhos: Lira, Trapézio Fixo e Tecido.',
  'Circo - Acrobacia': 'A CONTRATADA ministra aulas regulares de acrobacia de solo, em grupo, com uma hora de duração, abrangendo preparação física, elementos acrobáticos, equilíbrios e progressões conforme o nível do(a) aluno(a).',
  'Circo Infantil': 'A CONTRATADA ministra aulas regulares de circo infantil, em grupo, destinadas a crianças na faixa etária de 4 a 10 anos, com uma hora de duração, abrangendo a iniciação às linguagens circenses (equilíbrio, acrobacia de solo, manipulação e aparelhos adaptados à idade), com abordagem lúdica e progressiva, sob supervisão técnica.',
};

// Cada modalidade tem uma lista de DECLARACOES individuais do Anexo I.
// Cada declaracao vira, no formulario, um checkbox proprio (nao um "li e concordo" unico),
// conforme a diretriz juridica: aceite item por item, com registro de auditoria granular.
const ANEXO_I_ITENS_POR_MODALIDADE = {
  'Yoga': (nome, rg, cpf) => [
    `Encontro-me em condições de saúde compatíveis com a prática de yoga e não possuo, até onde é de meu conhecimento, restrição médica que a impeça.`,
    `Comprometo-me a informar à CONTRATADA, de forma tempestiva, qualquer condição de saúde, lesão, gestação, hipertensão (relevante para posturas invertidas), cirurgia ou restrição que possa afetar minha prática segura, bem como quaisquer alterações supervenientes.`,
    `Comprometo-me a respeitar meus próprios limites físicos e a seguir as orientações técnicas transmitidas pelo(a) instrutor(a).`,
    `Estou ciente de que a contratação de seguro de acidentes pessoais é recomendada, a título informativo, ficando a sua adesão a meu exclusivo critério.`,
    `Autorizo o atendimento emergencial e o acionamento dos serviços de socorro em caso de necessidade durante as atividades.`,
  ],

  'Aéreos': (nome, rg, cpf) => [
    `Estou ciente de que as atividades de circo aéreo (Lira, Trapézio Fixo e Tecido) envolvem esforço físico intenso e riscos inerentes à prática, assumindo participar delas de forma consciente e voluntária.`,
    `Encontro-me em condições de saúde compatíveis com a prática das atividades e não possuo, até onde é de meu conhecimento, restrição médica que a impeça.`,
    `Comprometo-me a informar à CONTRATADA, de forma tempestiva, qualquer condição de saúde, lesão, gestação, cirurgia ou restrição que possa afetar minha prática segura, bem como quaisquer alterações supervenientes.`,
    `Comprometo-me a seguir as orientações técnicas e as normas de segurança transmitidas pelos(as) instrutores(as), utilizando os aparelhos apenas com autorização e acompanhamento adequados ao meu nível.`,
    `Estou ciente de que a contratação de seguro de acidentes pessoais é recomendada, a título informativo, ficando a sua adesão a meu exclusivo critério.`,
    `Autorizo o atendimento emergencial e o acionamento dos serviços de socorro em caso de necessidade durante as atividades.`,
  ],

  'Circo - Acrobacia': (nome, rg, cpf) => [
    `Estou ciente de que as atividades de acrobacia de solo envolvem esforço físico intenso e riscos inerentes à prática, assumindo participar delas de forma consciente e voluntária.`,
    `Encontro-me em condições de saúde compatíveis com a prática das atividades e não possuo, até onde é de meu conhecimento, restrição médica que a impeça.`,
    `Comprometo-me a informar à CONTRATADA, de forma tempestiva, qualquer condição de saúde, lesão, gestação, cirurgia ou restrição que possa afetar minha prática segura, bem como quaisquer alterações supervenientes.`,
    `Comprometo-me a seguir as orientações técnicas e as normas de segurança transmitidas pelos(as) instrutores(as), praticando os elementos acrobáticos apenas com autorização e acompanhamento adequados ao meu nível.`,
    `Estou ciente de que a contratação de seguro de acidentes pessoais é recomendada, a título informativo, ficando a sua adesão a meu exclusivo critério.`,
    `Autorizo o atendimento emergencial e o acionamento dos serviços de socorro em caso de necessidade durante as atividades.`,
  ],

  'Circo Infantil': (nomeResp, rgResp, cpfResp, nomeCrianca, dataNascCrianca) => [
    `A criança está em condições de saúde compatíveis com a prática das atividades de circo infantil.`,
    `Comprometo-me a informar à CONTRATADA qualquer condição, lesão ou restrição relevante à sua prática segura.`,
    `Estou ciente de que as atividades circenses envolvem esforço físico e riscos inerentes, adaptados à faixa etária e conduzidos sob supervisão técnica.`,
    `Autorizo o atendimento emergencial e o acionamento dos serviços de socorro em caso de necessidade.`,
  ],
};

function montarCabecalhoAnexoI(modalidade, nome, rg, cpf, nomeResp, rgResp, cpfResp, nomeCrianca, dataNascCrianca) {
  if (modalidade === 'Circo Infantil') {
    return `Eu, ${nomeResp}, RG ${rgResp}, CPF ${cpfResp}, na condição de responsável legal por ${nomeCrianca}, nascido(a) em ${dataNascCrianca}, declaro, para os devidos fins, mediante o aceite individual de cada item abaixo, que:`;
  }
  const atividade = modalidade === 'Yoga' ? 'yoga' : modalidade === 'Aéreos' ? 'circo aéreo' : 'acrobacia de solo';
  return `Eu, ${nome}, RG ${rg}, CPF ${cpf}, na condição de aluno(a) das atividades de ${atividade} ministradas pela LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS (Espaço Liquidificador), declaro, para os devidos fins, mediante o aceite individual de cada item abaixo, que:`;
}

function montarTextoContratoMatricula(dados) {
  const {
    modalidade, plano, frequencia, turmas, valorMensal,
    nome, rg, cpf, endereco,
    nomeResponsavel, rgResponsavel, cpfResponsavel, nomeCrianca, dataNascCrianca,
    dataValidade,
  } = dados;

  const ehInfantil = modalidade === 'Circo Infantil';
  const nomeContratante = ehInfantil ? nomeResponsavel : nome;
  const rgContratante = ehInfantil ? rgResponsavel : rg;
  const cpfContratante = ehInfantil ? cpfResponsavel : cpf;

  const objeto = OBJETO_POR_MODALIDADE[modalidade] || '';
  const diaHorario = turmas.join(' e ');

  return `CONTRATO DE PRESTAÇÃO DE SERVIÇOS
Modelo-base — aceite eletrônico

DAS PARTES

Pelo presente instrumento particular de contrato de prestação de serviços, a empresa sob denominação social de CRISTIANE SOCCI LEONEL ME – nome fantasia LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS, devidamente inscrita no CNPJ 28.398.119/0001-83, com sede nesta capital, na Rua Doutor Carvalho de Mendonça, 67 – Campos Elíseos – CEP 01201-010, São Paulo/SP, doravante denominada CONTRATADA; e, de outro lado, ${ehInfantil ? 'o(a) responsável legal' : 'o(a) aluno(a)'} ${nomeContratante}, portador(a) do RG ${rgContratante} e CPF ${cpfContratante}, domiciliado(a) à ${endereco}, doravante denominado(a) CONTRATANTE${ehInfantil ? `, na condição de responsável por ${nomeCrianca}` : ''}, têm entre si, livremente ajustado e acordado, o seguinte:

CLÁUSULA PRIMEIRA – DO OBJETO

${objeto}

O(A) CONTRATANTE opta, por livre escolha, pelo seguinte plano:

${plano}, ${frequencia} – ${diaHorario}.

Totalizando o valor mensal de R$ ${valorMensal.toFixed(2)}, a ser pago até o 10º (décimo) dia do mês vigente.

Validade: a partir de ${dataValidade}.

CLÁUSULA SEGUNDA – CONDIÇÕES DE FÉRIAS / SUSPENSÃO DE PLANO

No plano semestral, a(o) aluna(o) poderá solicitar a suspensão de 15 (quinze) dias de férias e, no plano anual, 30 (trinta) dias de férias. Este período será acrescentado ao final do plano.

CLÁUSULA TERCEIRA – VALORES E FORMAS DE PAGAMENTO

Os pagamentos deverão ser realizados via PIX ou transferência bancária. Em caso de ingresso do(a) aluno(a) no decorrer do mês, será calculado valor proporcional no mês de início. Haverá correção anual dos valores, acompanhando os índices inflacionários. Quaisquer alterações serão comunicadas pela CONTRATADA com a antecedência necessária.

CLÁUSULA QUARTA – VENCIMENTOS E ATRASOS

A data estabelecida para pagamento é até o dia 10 (dez) do mês vigente. Em caso de não pagamento no vencimento, o valor será acrescido de multa de 2% (dois por cento) e juros de mora de 0,33% (trinta e três centésimos por cento) ao dia, até a data da efetiva quitação. Persistindo o atraso, a CONTRATADA encaminhará a cobrança a escritório especializado.

CLÁUSULA QUINTA – DA VIGÊNCIA E DA RENOVAÇÃO

O plano contratado vigora pelo período correspondente à sua modalidade — mensal (1 mês), semestral (6 meses) ou anual (12 meses) — contado a partir da data de validade indicada na Cláusula Primeira.

Ao final do período, o plano renova-se automaticamente, mantidas as condições e o valor vigentes, até que seja cancelado por qualquer das partes na forma da Cláusula Sexta, não sendo necessária qualquer confirmação periódica de continuidade por parte do(a) CONTRATANTE.

Parágrafo único – o período de fidelidade corresponde ao período inicialmente contratado. Cumprido esse período, a renovação prossegue sem novo prazo de fidelidade, podendo o(a) CONTRATANTE cancelar a qualquer tempo, mediante o aviso prévio da Cláusula Sexta, sem incidência de multa.

CLÁUSULA SEXTA – DO CANCELAMENTO E DESLIGAMENTO

6.1. O cancelamento por iniciativa do(a) CONTRATANTE deverá ser comunicado por escrito, com 30 (trinta) dias de antecedência. As aulas e as cobranças permanecem durante esse período de aviso, encerrando-se ao seu término.

6.2. Plano mensal: não há período de fidelidade nem multa; o cancelamento opera-se apenas com o aviso prévio de 30 (trinta) dias.

6.3. Planos semestral e anual: o desconto concedido em relação ao plano mensal é condicionado à permanência pelo período contratado. Em caso de cancelamento antes do término do período de fidelidade, o(a) CONTRATANTE perde o desconto relativo aos meses efetivamente cursados, pagando, a título de multa de cancelamento, a diferença entre o valor do plano mensal de mesma frequência e o valor do plano contratado, multiplicada pelo número de meses cursados:

Multa = (valor mensal − valor do plano contratado) × número de meses cursados.

6.4. O encerramento ao final do período contratado, ou após cumprido o período de fidelidade, não gera multa, desde que observado o aviso prévio de 30 (trinta) dias.

6.5. O abandono das aulas sem aviso prévio não caracteriza cancelamento; as cobranças permanecem até a efetiva comunicação por escrito do(a) CONTRATANTE, a partir da qual passa a correr o aviso prévio de 30 (trinta) dias.

6.6. A CONTRATADA dará ciência, por escrito, do recebimento da comunicação de cancelamento, disponibilizando a vaga aos demais interessados ao término do vínculo.

CLÁUSULA SÉTIMA – FERIADOS, RECESSOS, FALTAS E REPOSIÇÕES

Não haverá reposição de aula em feriados nacionais, estaduais e municipais. Haverá recesso de fim de ano, conforme calendário informado pela CONTRATADA. A reposição de faltas será possível mediante solicitação por escrito com 2 (dois) dias úteis de antecedência, sujeita à disponibilidade de vagas. Máximo de 1 (uma) reposição mensal para 1x/semana, ou 2 (duas) para 2x/semana.

CLÁUSULA OITAVA – SAÚDE, RESPONSABILIDADE E ASSUNÇÃO DE RISCO

${ehInfantil ? 'As atividades envolvem esforço físico e riscos inerentes, adaptados à faixa etária. O(A) responsável legal declara que a criança está em condições de saúde compatíveis com a atividade, conforme o Termo de Responsabilidade constante do Anexo I.' : 'As atividades envolvem esforço físico e riscos inerentes à prática. O(A) CONTRATANTE declara estar em condições de saúde compatíveis com a atividade, conforme o Termo de Responsabilidade e Declaração de Aptidão constante do Anexo I.'} O(A) CONTRATANTE compromete-se a seguir as orientações dos(as) instrutores(as) e as normas de segurança do espaço.

CLÁUSULA NONA – DA PROTEÇÃO DE DADOS PESSOAIS (LGPD)

Para a execução deste contrato, a CONTRATADA coleta e trata dados pessoais do(a) CONTRATANTE${ehInfantil ? ' e da criança' : ''}, em conformidade com a Lei nº 13.709/2018 (LGPD), exclusivamente para cadastro, gestão do plano, comunicação, cobrança e prática segura das atividades.${ehInfantil ? ' O tratamento de dados da criança observa o art. 14 da LGPD, mediante o consentimento específico e destacado do responsável legal.' : ''} Os dados não serão comercializados nem compartilhados com terceiros, salvo obrigação legal. O(A) titular pode, a qualquer tempo, solicitar acesso, correção ou eliminação de dados.

CLÁUSULA DÉCIMA – DO FORO

As partes elegem o foro da Comarca da Capital do Estado de São Paulo, facultado ao(à) CONTRATANTE optar pelo foro de seu domicílio, nos termos do CDC.

CLÁUSULA DÉCIMA PRIMEIRA – DO ACEITE ELETRÔNICO

Este contrato e seu Anexo I são celebrados por meio eletrônico. O(A) CONTRATANTE declara ter tido acesso ao inteiro teor deste instrumento e de cada declaração do Anexo I antes de manifestar sua concordância, mediante aceite eletrônico individual de cada item na ficha de inscrição, nos termos do art. 107 do Código Civil e da MP 2.200-2/2001. A autoria e a data são comprovadas pelo registro eletrônico que integra este contrato como prova de aceitação.`;
}

function montarAnexoIItens(dados) {
  const { modalidade, nome, rg, cpf, nomeResponsavel, rgResponsavel, cpfResponsavel, nomeCrianca, dataNascCrianca } = dados;
  const ehInfantil = modalidade === 'Circo Infantil';
  const cabecalho = montarCabecalhoAnexoI(modalidade, nome, rg, cpf, nomeResponsavel, rgResponsavel, cpfResponsavel, nomeCrianca, dataNascCrianca);
  const itens = ehInfantil
    ? ANEXO_I_ITENS_POR_MODALIDADE['Circo Infantil'](nomeResponsavel, rgResponsavel, cpfResponsavel, nomeCrianca, dataNascCrianca)
    : ANEXO_I_ITENS_POR_MODALIDADE[modalidade](nome, rg, cpf);
  return { cabecalho, itens };
}

// Quebra um texto longo em blocos de ate `tamanho` caracteres, no formato aceito por propriedades
// rich_text do Notion (cada bloco de texto tem limite de 2000 caracteres na API).
function paraRichTextEmBlocos(texto, tamanho = 1900) {
  const blocos = [];
  for (let i = 0; i < texto.length; i += tamanho) blocos.push({ text: { content: texto.slice(i, i + tamanho) } });
  return blocos.length ? blocos : [{ text: { content: '' } }];
}

// Duracao (em meses) do periodo de fidelidade de cada plano, conforme Clausula Quinta.
const DURACAO_FIDELIDADE_MESES = { Mensal: 1, Semestral: 6, Anual: 12 };

// Conta quantos ciclos mensais de vigencia ja comecaram entre dataInicio e dataReferencia (hoje, por padrao).
// Ex.: no proprio dia da matricula ja conta 1 (1o mes em curso); ao completar 2 meses corridos, conta 3
// (pois o 3o mes de vigencia ja comecou) — mesmo que incompleto, o mes cursado ja gerou o desconto de fidelidade.
function contarMesesCursados(dataInicioISO, dataReferencia) {
  if (!dataInicioISO) return 0;
  const inicio = new Date(dataInicioISO);
  const ref = dataReferencia || new Date();
  if (isNaN(inicio.getTime())) return 0;
  let meses = 0;
  const cursor = new Date(inicio.getTime());
  while (cursor <= ref) {
    meses++;
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return meses;
}

// Data em que o periodo atualmente contratado termina (inicio + duracao do plano), usada para
// popular "Vencimento do Contrato" no Notion (referencia de renovacao, nao usada no calculo da multa).
function calcularVencimentoContrato(dataInicioISO, plano) {
  if (!dataInicioISO) return null;
  const inicio = new Date(dataInicioISO);
  if (isNaN(inicio.getTime())) return null;
  const duracao = DURACAO_FIDELIDADE_MESES[plano] || 1;
  inicio.setMonth(inicio.getMonth() + duracao);
  return inicio.toISOString();
}

// ---------------------------------------------------------------------------
// FERIAS / SUSPENSAO DO PLANO (Portal Aluna)
// ---------------------------------------------------------------------------
// Dias de ferias por ciclo do plano (Clausula Segunda). Mensal nao tem direito.
const FERIAS_DIAS_POR_PLANO = { Mensal: 0, Semestral: 15, Anual: 30 };
// Antecedencia minima, em dias, para solicitar um periodo de ferias.
const FERIAS_AVISO_MINIMO_DIAS = 7;

// Zera a hora de uma data, para comparar apenas dia/mes/ano.
function apenasData(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

// Quantidade de dias corridos entre duas datas ISO (inclusive nas duas pontas).
function diasNoIntervaloFerias(inicioISO, fimISO) {
  const inicio = apenasData(inicioISO);
  const fim = apenasData(fimISO);
  return Math.round((fim.getTime() - inicio.getTime()) / 86400000) + 1;
}

// Janela do ciclo atual do plano (inicio/fim) que contem dataReferencia. Caminha ciclo a
// ciclo a partir da data de inicio da matricula; cada ciclo tem seu fim nominal (inicio +
// duracao do plano) estendido pelos dias de ferias tirados DENTRO daquele ciclo, e o ciclo
// seguinte comeca exatamente onde o anterior terminou (ja com a extensao). Assim o calculo
// acompanha a extensao real do plano, sem depender do campo "Vencimento do Contrato".
function calcularCicloAtualFerias(dataInicioISO, plano, periodosFerias, dataReferencia) {
  const duracao = DURACAO_FIDELIDADE_MESES[plano] || 1;
  const periodos = periodosFerias || [];
  const ref = dataReferencia || new Date();
  let cursor = new Date(dataInicioISO);
  while (true) {
    const nominalFim = new Date(cursor.getTime());
    nominalFim.setMonth(nominalFim.getMonth() + duracao);
    const diasFeriasNesteCiclo = periodos.reduce((total, p) => {
      const inicioPeriodo = new Date(p.inicio);
      if (inicioPeriodo >= cursor && inicioPeriodo < nominalFim) {
        return total + diasNoIntervaloFerias(p.inicio, p.fim);
      }
      return total;
    }, 0);
    const fimReal = new Date(nominalFim.getTime());
    fimReal.setDate(fimReal.getDate() + diasFeriasNesteCiclo);
    if (fimReal > ref) {
      return { inicioCiclo: cursor, fimCiclo: fimReal };
    }
    cursor = fimReal;
  }
}

// Extrai os periodos de ferias ja registrados no texto de Observacoes da matricula
// (linhas no formato "FERIAS: AAAA-MM-DD a AAAA-MM-DD (...)").
function extrairPeriodosFerias(observacoes) {
  if (!observacoes) return [];
  const regex = /FERIAS:\s*(\d{4}-\d{2}-\d{2})\s*a\s*(\d{4}-\d{2}-\d{2})/g;
  const periodos = [];
  let m;
  while ((m = regex.exec(observacoes)) !== null) {
    periodos.push({ inicio: m[1], fim: m[2] });
  }
  return periodos;
}

// Soma os dias dos periodos de ferias cujo inicio cai dentro do ciclo informado.
function diasFeriasUsadosNoCiclo(periodos, inicioCiclo, fimCiclo) {
  return periodos.reduce((total, p) => {
    const inicioPeriodo = new Date(p.inicio);
    if (inicioPeriodo >= inicioCiclo && inicioPeriodo < fimCiclo) {
      return total + diasNoIntervaloFerias(p.inicio, p.fim);
    }
    return total;
  }, 0);
}

// Verifica se um novo periodo [inicio,fim] se sobrepoe a algum periodo ja registrado.
function periodoFeriasSobrepoe(periodos, inicioISO, fimISO) {
  const novoInicio = apenasData(inicioISO);
  const novoFim = apenasData(fimISO);
  return periodos.some(p => {
    const pInicio = apenasData(p.inicio);
    const pFim = apenasData(p.fim);
    return novoInicio <= pFim && pInicio <= novoFim;
  });
}

// Verifica se a data informada (AAAA-MM-DD) cai dentro de algum periodo de ferias.
function dataEstaEmFerias(periodos, dataISO) {
  if (!dataISO) return false;
  const alvo = apenasData(dataISO);
  return periodos.some(p => alvo >= apenasData(p.inicio) && alvo <= apenasData(p.fim));
}


// Multa de cancelamento (Clausula Sexta, item 6.3): so incide em planos semestral/anual, e so antes de
// cumprida a fidelidade. Formula (regra final, substitui a antiga baseada em desconto perdido):
//   multa = minimo( valor_plano x meses_restantes x 0,30 ; valor_plano x 2 )
// meses_restantes = meses que faltam para completar a fidelidade, contados a partir do FIM do aviso
// previo de 30 dias (nao da data de hoje) — Clausula 6.3.1. Teto: 2 mensalidades do plano contratado.
// Reutilizavel: chamada tanto pelo Portal Aluna (simular/solicitar cancelamento) quanto, futuramente, pelo Portal Admin.
const AVISO_PREVIO_DIAS_CANCELAMENTO = 30;

function calcularMultaCancelamento(modalidade, frequencia, plano, dataInicio, dataReferencia) {
  const dadosModalidade = MODALIDADES_MATRICULA[modalidade];
  if (!dadosModalidade) return { ok: false, erro: 'Modalidade inválida.' };
  const precoTabela = dadosModalidade.precos[frequencia];
  if (!precoTabela) return { ok: false, erro: 'Frequência inválida para esta modalidade.' };
  const valorPlanoContratado = precoTabela[plano];
  if (valorPlanoContratado === undefined) return { ok: false, erro: 'Plano inválido para esta frequência/modalidade.' };

  const duracaoFidelidadeMeses = DURACAO_FIDELIDADE_MESES[plano] || 1;

  // Fim do aviso previo: referencia (hoje, por padrao) + 30 dias. E a partir dessa data que se contam
  // os meses restantes de fidelidade (Clausula 6.3.1).
  const refBase = dataReferencia || new Date();
  const dataFimAvisoPrevio = new Date(refBase.getTime());
  dataFimAvisoPrevio.setDate(dataFimAvisoPrevio.getDate() + AVISO_PREVIO_DIAS_CANCELAMENTO);

  const mesesCursados = Math.min(contarMesesCursados(dataInicio, dataFimAvisoPrevio), duracaoFidelidadeMeses);
  const mesesRestantes = Math.max(0, duracaoFidelidadeMeses - mesesCursados);

  if (plano === 'Mensal') {
    return {
      ok: true, multa: 0, mesesCursados, mesesRestantes: 0, duracaoFidelidadeMeses, fidelidadeCumprida: true,
      valorPlanoContratado, tetoAplicado: false,
      motivo: 'Plano mensal não possui período de fidelidade nem multa de cancelamento.',
    };
  }

  const fidelidadeCumprida = mesesRestantes <= 0;
  const multaSemTeto = valorPlanoContratado * mesesRestantes * 0.30;
  const tetoMulta = valorPlanoContratado * 2;
  const tetoAplicado = !fidelidadeCumprida && multaSemTeto > tetoMulta;
  const multa = fidelidadeCumprida ? 0 : Math.round(Math.min(multaSemTeto, tetoMulta) * 100) / 100;

  return {
    ok: true, multa, mesesCursados, mesesRestantes, duracaoFidelidadeMeses, fidelidadeCumprida,
    valorPlanoContratado, tetoAplicado,
    motivo: fidelidadeCumprida ? 'Período de fidelidade já cumprido — cancelamento sem multa.' : null,
  };
}

async function gerarPdfContratoMatricula(textoContrato, { assinaturaDigitada, dataHoraISO, ip, dispositivo, versao }) {
  const PDFDocument = require('pdfkit');
  const chunks = [];
  const doc = new PDFDocument({ margin: 50 });
  doc.on('data', (c) => chunks.push(c));
  const fimPromise = new Promise((resolve) => doc.on('end', () => resolve(Buffer.concat(chunks))));

  doc.fontSize(9).text(textoContrato, { align: 'left' });
  doc.moveDown(2);
  doc.fontSize(9).text('--- REGISTRO DE ACEITE ELETRÔNICO ---');
  doc.text('Assinado por: ' + assinaturaDigitada);
  doc.text('Data/Hora: ' + dataHoraISO);
  doc.text('IP: ' + ip);
  doc.text('Dispositivo: ' + dispositivo);
  doc.text('Versão do documento: ' + versao);

  doc.end();
  return await fimPromise;
}

app.post('/matricula/contrato-texto', (req, res) => {
  try {
    const texto = montarTextoContratoMatricula(req.body);
    const anexo = montarAnexoIItens(req.body);
    const crypto = require('crypto');
    const versao = crypto.createHash('sha256').update(texto + JSON.stringify(anexo)).digest('hex').slice(0, 16);
    res.json({ ok: true, texto, anexoCabecalho: anexo.cabecalho, anexoItens: anexo.itens, versao });
  } catch (err) {
    console.error('[matricula] erro ao montar contrato:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.get('/modalidades-matricula', (req, res) => {
  const resumo = {};
  for (const [nome, dados] of Object.entries(MODALIDADES_MATRICULA)) {
    resumo[nome] = { turmas: dados.turmas.map(t => t.nome), permiteFrequenciaDupla: dados.permiteFrequenciaDupla, precos: dados.precos };
  }
  res.json({ ok: true, modalidades: resumo });
});

app.get('/aluno/:cpf', async (req, res) => {
  const cpfLimpo = req.params.cpf.replace(/\D/g, '');
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'CPF', rich_text: { equals: cpfLimpo } }, page_size: 5 }),
    });
    const d = await r.json();
    const resultados = d.results || [];
    if (resultados.length === 0) return res.json({ ok: true, encontrado: false });

    const p = resultados[0].properties;
    res.json({
      ok: true, encontrado: true,
      nome: p['Nome']?.title?.[0]?.plain_text || '',
      contato: p['Contato']?.phone_number || '',
      contatoEmergenciaNome: p['Contato de Emergência']?.rich_text?.[0]?.plain_text || '',
      contatoEmergenciaTelefone: p['Tel. Emergência']?.phone_number || '',
      rg: p['RG']?.rich_text?.[0]?.plain_text || '',
      endereco: p['Endereço']?.rich_text?.[0]?.plain_text || '',
      email: p['Email']?.email || '',
      dataNascimento: p['Data de Nascimento']?.date?.start || '',
    });
  } catch (err) {
    console.error('[aluno] erro ao buscar CPF:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

async function buscarCapacidadeDaTurma(modalidade) {
  const r = await fetch('https://api.notion.com/v1/databases/' + CAPACIDADE_TURMAS_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { property: 'Modalidade', select: { equals: modalidade } },
      page_size: 100,
    }),
  });
  const d = await r.json();
  const mapa = {};
  (d.results || []).forEach(pagina => {
    const p = pagina.properties;
    const turmaNome = p['Título']?.title?.[0]?.plain_text || '';
    const capacidade = p['Capacidade Regular']?.number;
    if (turmaNome && capacidade != null) mapa[turmaNome] = capacidade;
  });
  return mapa;
}

// Busca o(s) professor(es) de cada turma de uma modalidade no banco "Capacidade de
// Turmas" (campo multi_select "Professor" — permite mais de um nome por turma, ex.:
// substitutos ou testes). Usada por turmasDoProfessor() para nao depender mais so do
// valor fixo em MODALIDADES_MATRICULA — permite trocar professor de turma direto no
// Notion, sem precisar de patch no server.js.
async function buscarProfessorDaTurma(modalidade) {
  const r = await fetch('https://api.notion.com/v1/databases/' + CAPACIDADE_TURMAS_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { property: 'Modalidade', select: { equals: modalidade } },
      page_size: 100,
    }),
  });
  const d = await r.json();
  const mapa = {};
  (d.results || []).forEach(pagina => {
    const p = pagina.properties;
    const turmaNome = p['Título']?.title?.[0]?.plain_text || '';
    const professores = (p['Professor']?.multi_select || []).map(opt => opt.name).filter(Boolean);
    if (turmaNome && professores.length > 0) mapa[turmaNome] = professores;
  });
  return mapa;
}

app.get('/vagas-modalidade/:modalidade', async (req, res) => {
  const modalidade = decodeURIComponent(req.params.modalidade);
  const dadosModalidade = MODALIDADES_MATRICULA[modalidade];
  if (!dadosModalidade) return res.status(400).json({ ok: false, erro: 'Modalidade não encontrada.' });

  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Modalidade', select: { equals: modalidade } },
          { property: 'Status', select: { equals: 'Ativa' } },
        ]},
        page_size: 200,
      }),
    });
    const d = await r.json();
    const ocupacaoPorTurma = {};
    (d.results || []).forEach(page => {
      const turma = page.properties['Turma']?.select?.name;
      if (turma) ocupacaoPorTurma[turma] = (ocupacaoPorTurma[turma] || 0) + 1;
    });

    const capacidadeReal = await buscarCapacidadeDaTurma(modalidade);
    const turmas = dadosModalidade.turmas.map(t => {
      const ocupadas = ocupacaoPorTurma[t.nome] || 0;
      const limiteReal = capacidadeReal[t.nome] ?? t.limite;
      return { ...t, limite: limiteReal, ocupadas, vagasRestantes: Math.max(0, limiteReal - ocupadas) };
    });

    res.json({ ok: true, modalidade, turmas, precos: dadosModalidade.precos, permiteFrequenciaDupla: dadosModalidade.permiteFrequenciaDupla });
  } catch (err) {
    console.error('[vagas-modalidade] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/matricula/inscrever', async (req, res) => {
  const {
    nome, cpf, rg, endereco, email, dataNascimento, contato, contatoEmergenciaNome, contatoEmergenciaTelefone,
    possuiAlergias, quaisAlergias, usaMedicamentos, quaisMedicamentos,
    condicaoSaude, qualCondicao, cirurgiasLesoes, detalhesCirurgias, liberadaAtividadeFisica,
    modalidade, turmas, frequencia, plano, observacoes,
    nomeResponsavel, rgResponsavel, cpfResponsavel, parentesco, autorizadosRetirar,
    consentimentoDadosPessoais, consentimentoDadosSaude, consentimentoContrato, consentimentoUsoImagem,
    anexoAceites,
    assinaturaDigitada, versaoContrato, dispositivo,
  } = req.body;
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

  const ehInfantil = modalidade === 'Circo Infantil';

  if (!cpf || !contato || !contatoEmergenciaNome || !contatoEmergenciaTelefone ||
      !modalidade || !turmas || !turmas.length || !frequencia || !plano ||
      !endereco || !email || !dataNascimento || !assinaturaDigitada ||
      !consentimentoDadosPessoais || !consentimentoDadosSaude || !consentimentoContrato) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  if (!Array.isArray(anexoAceites) || anexoAceites.length === 0 || anexoAceites.some(v => !v)) {
    const itensFaltantes = Array.isArray(anexoAceites) ? anexoAceites.map((v, i) => (!v ? i + 1 : null)).filter(Boolean) : [];
    return res.status(400).json({
      ok: false,
      erro: itensFaltantes.length
        ? 'Você precisa marcar todos os itens do Anexo I (Termo de Responsabilidade). Item(ns) pendente(s): ' + itensFaltantes.join(', ') + '.'
        : 'Você precisa marcar todos os itens do Anexo I (Termo de Responsabilidade).',
      itensFaltantes,
    });
  }
  if (ehInfantil && (!nomeResponsavel || !rgResponsavel || !cpfResponsavel || !nome)) {
    return res.status(400).json({ ok: false, erro: 'Preencha os dados do responsável legal e da criança.' });
  }
  if (!ehInfantil && (!nome || !rg)) {
    return res.status(400).json({ ok: false, erro: 'Preencha nome e RG.' });
  }

  const dadosModalidade = MODALIDADES_MATRICULA[modalidade];
  if (!dadosModalidade) return res.status(400).json({ ok: false, erro: 'Modalidade inválida.' });
  const precoTabela = dadosModalidade.precos[frequencia];
  if (!precoTabela || !precoTabela[plano]) return res.status(400).json({ ok: false, erro: 'Combinação de frequência/plano inválida.' });
  const valorTotal = precoTabela[plano];
  const valorPorTurma = Math.round((valorTotal / turmas.length) * 100) / 100;

  try {
    // Revalida vagas em tempo real antes de gravar (evita corrida entre duas inscricoes simultaneas)
    const rCheck = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Modalidade', select: { equals: modalidade } },
          { property: 'Status', select: { equals: 'Ativa' } },
        ]},
        page_size: 200,
      }),
    });
    const dCheck = await rCheck.json();
    const ocupacaoAtual = {};
    (dCheck.results || []).forEach(page => {
      const t = page.properties['Turma']?.select?.name;
      if (t) ocupacaoAtual[t] = (ocupacaoAtual[t] || 0) + 1;
    });
    for (const nomeTurma of turmas) {
      const infoTurma = dadosModalidade.turmas.find(t => t.nome === nomeTurma);
      if (!infoTurma) return res.json({ ok: false, erro: 'Turma inválida: ' + nomeTurma });
      const ocupadas = ocupacaoAtual[nomeTurma] || 0;
      if (ocupadas >= infoTurma.limite) {
        return res.json({ ok: false, erro: 'A turma ' + nomeTurma + ' acabou de lotar. Escolha outro horário.' });
      }
    }

    // Monta e assina o contrato uma unica vez (vale para todas as turmas desta matricula)
    const dataHoraAceiteISO = new Date().toISOString();
    const dataValidadeFmt = new Date().toLocaleDateString('pt-BR');
    const dataNascCriancaFmt = ehInfantil ? new Date(dataNascimento).toLocaleDateString('pt-BR') : '';
    const dadosContratoBase = {
      modalidade, plano, frequencia, turmas, valorMensal: valorTotal,
      nome, rg, cpf, endereco,
      nomeResponsavel, rgResponsavel, cpfResponsavel,
      nomeCrianca: ehInfantil ? nome : '', dataNascCrianca: dataNascCriancaFmt,
      dataValidade: dataValidadeFmt,
    };
    const textoCorpo = montarTextoContratoMatricula(dadosContratoBase);
    const anexo = montarAnexoIItens(dadosContratoBase);
    const anexoTextoNumerado = anexo.itens.map((item, i) => (i + 1) + '. ' + item).join('\n');
    const textoContrato = textoCorpo + '\n\nANEXO I — TERMO DE RESPONSABILIDADE E DECLARAÇÃO DE APTIDÃO\n\n' + anexo.cabecalho + '\n\n' + anexoTextoNumerado;
    const crypto = require('crypto');
    const versaoCalculada = versaoContrato || crypto.createHash('sha256').update(textoContrato).digest('hex').slice(0, 16);

    // Trilha de auditoria item a item do Anexo I: cada declaracao marcada gera seu proprio registro
    // (identificador do item, versao do texto do contrato, marcado sim/nao, data/hora, IP), conforme diretriz.
    const anexoTrilhaAuditoria = anexo.itens.map((item, i) =>
      (i + 1) + '. ' + item +
      '\n   [item ' + (i + 1) + '/' + anexo.itens.length + ' | aceito=' + (anexoAceites[i] ? 'SIM' : 'NÃO') +
      ' | versão=' + versaoCalculada + ' | ' + dataHoraAceiteISO + ' | IP ' + ip + ']'
    ).join('\n');

    let linkContratoPdf = '';
    try {
      const pdfBuffer = await gerarPdfContratoMatricula(textoContrato, {
        assinaturaDigitada, dataHoraISO: dataHoraAceiteISO, ip, dispositivo: dispositivo || '', versao: versaoCalculada,
      });
      const msToken = await getMicrosoftToken();
      const nomePastaContrato = 'Matriculas-' + slugify(modalidade) + '-' + slugify(ehInfantil ? nomeResponsavel : nome);
      const folderId = await criarOuObterSubpasta(msToken, nomePastaContrato);
      const pdfUploaded = await uploadBufferOneDrive(msToken, folderId, pdfBuffer, 'contrato-' + Date.now() + '.pdf');
      linkContratoPdf = await criarLinkCompartilhamento(msToken, pdfUploaded.id);
    } catch (e) {
      console.error('[matricula] erro ao gerar/subir PDF do contrato:', e.message);
    }

    for (const nomeTurma of turmas) {
      const infoTurma = dadosModalidade.turmas.find(t => t.nome === nomeTurma);
      await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: ALUNAS_DB },
          properties: {
            'Nome': { title: [{ text: { content: nome } }] },
            'CPF': { rich_text: [{ text: { content: (ehInfantil ? cpfResponsavel : cpf).replace(/\D/g, '') } }] },
            'RG': { rich_text: [{ text: { content: rg || '' } }] },
            'Data de Nascimento': { date: { start: dataNascimento } },
            'Endereço': { rich_text: [{ text: { content: endereco } }] },
            'Email': { email: email },
            'Contato': { phone_number: contato },
            'Contato de Emergência': { rich_text: [{ text: { content: contatoEmergenciaNome } }] },
            'Tel. Emergência': { phone_number: contatoEmergenciaTelefone },
            'Possui alergias?': { select: { name: possuiAlergias || 'Não' } },
            'Quais alergias?': { rich_text: [{ text: { content: quaisAlergias || '' } }] },
            'Usa medicamentos?': { select: { name: usaMedicamentos || 'Não' } },
            'Quais medicamentos?': { rich_text: [{ text: { content: quaisMedicamentos || '' } }] },
            'Condição de saúde?': { select: { name: condicaoSaude || 'Não' } },
            'Qual condição?': { rich_text: [{ text: { content: qualCondicao || '' } }] },
            'Cirurgias ou lesões?': { select: { name: cirurgiasLesoes || 'Não' } },
            'Detalhes cirurgias/lesões': { rich_text: [{ text: { content: detalhesCirurgias || '' } }] },
            'Liberada p/ atividade física?': { select: { name: liberadaAtividadeFisica || 'Sim' } },
            'Modalidade': { select: { name: modalidade } },
            'Turma': { select: { name: nomeTurma } },
            'Professor': { select: { name: infoTurma.professor } },
            'Dia': { select: { name: infoTurma.dia } },
            'Horário': { select: { name: infoTurma.horario } },
            'Frequência': { select: { name: frequencia } },
            'Plano': { select: { name: plano } },
            'Valor': { number: valorPorTurma },
            'Status': { select: { name: 'Ativa' } },
            'Observações': { rich_text: [{ text: { content: observacoes || '' } }] },
            'Nome do Responsável': { rich_text: [{ text: { content: ehInfantil ? nomeResponsavel : '' } }] },
            'RG do Responsável': { rich_text: [{ text: { content: ehInfantil ? rgResponsavel : '' } }] },
            'Parentesco': { rich_text: [{ text: { content: ehInfantil ? (parentesco || '') : '' } }] },
            'Data de Nascimento da Criança': ehInfantil ? { date: { start: dataNascimento } } : { date: null },
            'Autorizados a Retirar': { rich_text: [{ text: { content: ehInfantil ? (autorizadosRetirar || '') : '' } }] },
            'Consentimento Dados Pessoais': { checkbox: !!consentimentoDadosPessoais },
            'Consentimento Dados de Saúde': { checkbox: !!consentimentoDadosSaude },
            'Consentimento Uso de Imagem': { checkbox: !!consentimentoUsoImagem },
            'IP Aceite': { rich_text: [{ text: { content: ip } }] },
            'User-Agent Aceite': { rich_text: [{ text: { content: dispositivo || '' } }] },
            'Data/Hora Aceite Contrato': { date: { start: dataHoraAceiteISO } },
            'Versão do Contrato': { rich_text: [{ text: { content: versaoCalculada } }] },
            'Link do Contrato PDF': { url: linkContratoPdf || null },
            'Anexo I - Itens Aceitos': { rich_text: paraRichTextEmBlocos(anexoTrilhaAuditoria) },
            'Vencimento do Contrato': { date: { start: calcularVencimentoContrato(dataHoraAceiteISO, plano) } },
          },
        }),
      });
    }

    const primeiroNome = nome.split(' ')[0];
    const numLimpo = contato.replace(/\D/g, '');
    const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
    const turmasTexto = turmas.join(' e ');
    if (numBr) {
      try {
        await enviarWhatsApp(numBr, 'Olá, ' + primeiroNome + '! 🎉\n\nSua matrícula em *' + modalidade + '* (' + turmasTexto + ') foi registrada!\n\n💳 Plano ' + plano + ': R$ ' + valorTotal.toFixed(2) + '\n\nPara confirmar, faça o pagamento via Pix para:\nfabio@cialiquidificador.com.br\n\nMande o comprovante aqui pelo WhatsApp. Qualquer dúvida, é só chamar!');
      } catch(e) {}
    }
    const msgInterna = '🎉 *Nova matrícula* — ' + modalidade + '\nAluno(a): ' + nome + ' (' + contato + ')\nTurma(s): ' + turmasTexto + '\nPlano: ' + plano + ' (' + frequencia + ') — R$ ' + valorTotal.toFixed(2);
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}

    res.json({ ok: true, valorTotal });
  } catch (err) {
    console.error('[matricula] erro ao inscrever:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// MIGRAÇÃO DE CADASTRO — alunas com contrato antigo, ja ativas no sistema
// ============================================================
// Completa os dados que faltam (RG, endereco, saude, emergencia, consentimentos,
// aceite do Anexo I) e gera o PDF assinado, SEM criar pagina nova e SEM checar vaga
// (a aluna ja esta Ativa e ja ocupa a vaga). Nao mexe em "Data/Hora Aceite Contrato"
// nem em plano/turma/valor ja existentes — so completa o que falta.
app.post('/portal-aluna/completar-migracao', async (req, res) => {
  const {
    cpf, sessionToken, modalidade, turma, menorDeIdade,
    rg, endereco, email, dataNascimento, contatoEmergenciaNome, contatoEmergenciaTelefone,
    possuiAlergias, quaisAlergias, usaMedicamentos, quaisMedicamentos,
    condicaoSaude, qualCondicao, cirurgiasLesoes, detalhesCirurgias, liberadaAtividadeFisica,
    nomeResponsavel, rgResponsavel, cpfResponsavel, parentesco, autorizadosRetirar,
    consentimentoDadosPessoais, consentimentoDadosSaude, consentimentoUsoImagem,
    anexoAceites, assinaturaDigitada, dispositivo,
  } = req.body;
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

  const cpfLimpoMig = (cpf || '').replace(/\D/g, '');
  const dadosSessaoMig = verificarSessao(sessionToken);
  if (!dadosSessaoMig) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoMig.cpf !== cpfLimpoMig) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });

  if (!modalidade || !turma || !contatoEmergenciaNome || !contatoEmergenciaTelefone ||
      !endereco || !email || !dataNascimento || !assinaturaDigitada ||
      !consentimentoDadosPessoais || !consentimentoDadosSaude) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  if (!Array.isArray(anexoAceites) || anexoAceites.length === 0 || anexoAceites.some(v => !v)) {
    return res.status(400).json({ ok: false, erro: 'Você precisa marcar todos os itens do Anexo I (Termo de Responsabilidade).' });
  }

  const ehInfantil = modalidade === 'Circo Infantil';
  // representaResponsavel cobre tanto Circo Infantil quanto qualquer outra modalidade
  // em que a matricula seja de uma pessoa menor de idade (flag enviada pelo formulario).
  const representaResponsavel = ehInfantil || !!menorDeIdade;
  if (representaResponsavel && (!nomeResponsavel || !rgResponsavel)) {
    return res.status(400).json({ ok: false, erro: 'Preencha os dados do responsável legal.' });
  }
  if (!representaResponsavel && !rg) {
    return res.status(400).json({ ok: false, erro: 'Preencha o RG.' });
  }

  try {
    const matriculasMig = await buscarMatriculasPorCpf(cpfLimpoMig);
    const alvoMig = localizarMatricula(matriculasMig, modalidade, turma)[0] || null;
    if (!alvoMig) return res.status(404).json({ ok: false, erro: 'Matrícula não encontrada.' });

    const dadosModalidadeMig = MODALIDADES_MATRICULA[modalidade];
    if (!dadosModalidadeMig) return res.status(400).json({ ok: false, erro: 'Modalidade inválida.' });

    const precoTabelaMig = dadosModalidadeMig.precos[alvoMig.frequencia];
    const valorTabela = precoTabelaMig ? precoTabelaMig[alvoMig.plano] : undefined;
    const valorFinal = (alvoMig.valor && alvoMig.valor > 0) ? alvoMig.valor : valorTabela;
    if (valorFinal === undefined) {
      return res.status(400).json({ ok: false, erro: 'Não foi possível determinar o valor do plano — confira Plano/Frequência no Notion.' });
    }

    // CPF do responsavel: sempre o CPF usado no login (mesmo padrao ja usado para
    // Circo Infantil, onde o campo "CPF" da matricula guarda o CPF do responsavel).
    const cpfResponsavelFinal = representaResponsavel ? (cpfResponsavel || cpfLimpoMig) : cpfResponsavel;

    const ehMensal = alvoMig.plano === 'Mensal';
    let dataInicioReal = alvoMig.dataInicio;

    if (ehMensal) {
      dataInicioReal = dataInicioReal || new Date().toISOString();
    } else if (!dataInicioReal) {
      return res.status(400).json({
        ok: false,
        erro: 'Aguarde o Fábio preencher a data de início do seu contrato antes de continuar.',
      });
    }

    const agoraISO = new Date().toISOString();
    const dataValidadeFmt = new Date(dataInicioReal).toLocaleDateString('pt-BR');
    const dataNascCriancaFmt = representaResponsavel ? new Date(dataNascimento).toLocaleDateString('pt-BR') : '';
    const nomeParaContrato = alvoMig.nome;

    const dadosContratoBaseMig = {
      modalidade, plano: alvoMig.plano, frequencia: alvoMig.frequencia, turmas: [turma], valorMensal: valorFinal,
      nome: nomeParaContrato, rg, cpf: cpfLimpoMig, endereco,
      nomeResponsavel, rgResponsavel, cpfResponsavel: cpfResponsavelFinal,
      nomeCrianca: representaResponsavel ? nomeParaContrato : '', dataNascCrianca: dataNascCriancaFmt,
      dataValidade: dataValidadeFmt,
    };
    const textoCorpoMig = montarTextoContratoMatricula(dadosContratoBaseMig);
    const anexoMig = montarAnexoIItens(dadosContratoBaseMig);
    const anexoTextoNumeradoMig = anexoMig.itens.map((item, i) => (i + 1) + '. ' + item).join('\n');
    const textoContratoMig = textoCorpoMig + '\n\nANEXO I — TERMO DE RESPONSABILIDADE E DECLARAÇÃO DE APTIDÃO\n\n' + anexoMig.cabecalho + '\n\n' + anexoTextoNumeradoMig;
    const crypto = require('crypto');
    const versaoCalculadaMig = crypto.createHash('sha256').update(textoContratoMig).digest('hex').slice(0, 16);

    const anexoTrilhaAuditoriaMig = anexoMig.itens.map((item, i) =>
      (i + 1) + '. ' + item +
      '\n   [item ' + (i + 1) + '/' + anexoMig.itens.length + ' | aceito=' + (anexoAceites[i] ? 'SIM' : 'NÃO') +
      ' | versão=' + versaoCalculadaMig + ' | ' + agoraISO + ' | IP ' + ip + ']'
    ).join('\n');

    let linkContratoPdfMig = '';
    try {
      const pdfBufferMig = await gerarPdfContratoMatricula(textoContratoMig, {
        assinaturaDigitada, dataHoraISO: agoraISO, ip, dispositivo: dispositivo || '', versao: versaoCalculadaMig,
      });
      const msTokenMig = await getMicrosoftToken();
      const nomePastaContratoMig = 'Matriculas-' + slugify(modalidade) + '-' + slugify(nomeParaContrato);
      const folderIdMig = await criarOuObterSubpasta(msTokenMig, nomePastaContratoMig);
      const pdfUploadedMig = await uploadBufferOneDrive(msTokenMig, folderIdMig, pdfBufferMig, 'contrato-migracao-' + Date.now() + '.pdf');
      linkContratoPdfMig = await criarLinkCompartilhamento(msTokenMig, pdfUploadedMig.id);
    } catch (e) {
      console.error('[completar-migracao] erro ao gerar/subir PDF do contrato:', e.message);
    }

    const propsMigracao = {
      'RG': { rich_text: [{ text: { content: rg || '' } }] },
      'Endereço': { rich_text: [{ text: { content: endereco } }] },
      'Email': { email: email },
      'Data de Nascimento': { date: { start: dataNascimento } },
      'Contato de Emergência': { rich_text: [{ text: { content: contatoEmergenciaNome } }] },
      'Tel. Emergência': { phone_number: contatoEmergenciaTelefone },
      'Possui alergias?': { select: { name: possuiAlergias || 'Não' } },
      'Quais alergias?': { rich_text: [{ text: { content: quaisAlergias || '' } }] },
      'Usa medicamentos?': { select: { name: usaMedicamentos || 'Não' } },
      'Quais medicamentos?': { rich_text: [{ text: { content: quaisMedicamentos || '' } }] },
      'Condição de saúde?': { select: { name: condicaoSaude || 'Não' } },
      'Qual condição?': { rich_text: [{ text: { content: qualCondicao || '' } }] },
      'Cirurgias ou lesões?': { select: { name: cirurgiasLesoes || 'Não' } },
      'Detalhes cirurgias/lesões': { rich_text: [{ text: { content: detalhesCirurgias || '' } }] },
      'Liberada p/ atividade física?': { select: { name: liberadaAtividadeFisica || 'Sim' } },
      'Nome do Responsável': { rich_text: [{ text: { content: representaResponsavel ? nomeResponsavel : '' } }] },
      'RG do Responsável': { rich_text: [{ text: { content: representaResponsavel ? rgResponsavel : '' } }] },
      'Parentesco': { rich_text: [{ text: { content: representaResponsavel ? (parentesco || '') : '' } }] },
      'Autorizados a Retirar': { rich_text: [{ text: { content: representaResponsavel ? (autorizadosRetirar || '') : '' } }] },
      'Data de Nascimento da Criança': representaResponsavel ? { date: { start: dataNascimento } } : { date: null },
      'Consentimento Dados Pessoais': { checkbox: !!consentimentoDadosPessoais },
      'Consentimento Dados de Saúde': { checkbox: !!consentimentoDadosSaude },
      'Consentimento Uso de Imagem': { checkbox: !!consentimentoUsoImagem },
      'IP Aceite': { rich_text: [{ text: { content: ip } }] },
      'User-Agent Aceite': { rich_text: [{ text: { content: dispositivo || '' } }] },
      'Versão do Contrato': { rich_text: [{ text: { content: versaoCalculadaMig } }] },
      'Link do Contrato PDF': { url: linkContratoPdfMig || null },
      'Anexo I - Itens Aceitos': { rich_text: paraRichTextEmBlocos(anexoTrilhaAuditoriaMig) },
      'Data Aceite Migracao': { date: { start: agoraISO } },
    };
    if (!alvoMig.valor || alvoMig.valor <= 0) {
      propsMigracao['Valor'] = { number: valorFinal };
    }
    if (!alvoMig.vencimentoContrato) {
      propsMigracao['Vencimento do Contrato'] = { date: { start: calcularVencimentoContrato(dataInicioReal, alvoMig.plano) } };
    }
    if (ehMensal && !alvoMig.dataInicio) {
      propsMigracao['Data/Hora Aceite Contrato'] = { date: { start: dataInicioReal } };
    }

    const rPatchMig = await fetch('https://api.notion.com/v1/pages/' + alvoMig.pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: propsMigracao }),
    });
    if (!rPatchMig.ok) {
      const eBodyMig = await rPatchMig.text();
      console.error('[completar-migracao] Notion recusou atualizar matricula:', eBodyMig);
      return res.status(500).json({ ok: false, erro: 'Notion recusou gravar os dados — ver logs.' });
    }

    try {
      await enviarWhatsApp(
        WHATSAPP_FABIO,
        '📋 *Cadastro migrado* — ' + modalidade + '\nAluna: ' + nomeParaContrato +
        '\nTurma: ' + turma + '\nCadastro completo e contrato assinado digitalmente.'
      );
    } catch (eNotifMig) {
      console.error('[completar-migracao] erro ao notificar Fábio:', eNotifMig.message);
    }
    if (alvoMig.contato) {
      const numLimpoMig = alvoMig.contato.replace(/\D/g, '');
      const numBrMig = numLimpoMig.length === 11 ? '55' + numLimpoMig : numLimpoMig;
      if (numBrMig) {
        try {
          await enviarWhatsApp(numBrMig, 'Olá, ' + nomeParaContrato.split(' ')[0] + '! ✅\n\nSeu cadastro foi atualizado com sucesso e seu contrato digital foi gerado. Obrigado por completar as informações!');
        } catch (e) {}
      }
    }

    res.json({ ok: true, linkContratoPdf: linkContratoPdfMig });
  } catch (err) {
    console.error('[portal-aluna/completar-migracao] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// PORTAL PROFESSORES — Login por CPF, Minhas Turmas, Feriados
// ============================================================
const DECISOES_FERIADO_DB = 'ef0e4140-26a0-4829-bf68-7a24e4ba7618';

const PROFESSORES_PORTAL = [
  { nome: 'Gabi', cpf: '22946806855', telefone: '5511961416621' },
  { nome: 'Talita', cpf: '48458176831', telefone: '5511989142791' },
  { nome: 'Gustra', cpf: '45486626851', telefone: '5511988485740' },
  { nome: 'Guilherme', cpf: '', telefone: '5511989538880' },
  { nome: 'André', cpf: '36268818814', telefone: '5511981578744' },
  { nome: 'Titzi', cpf: '30814279830', telefone: '5511951780877' },
  { nome: 'Giulia', cpf: '51310549826', telefone: '5512988222584' },
  { nome: 'Roberta', cpf: '21811238882', telefone: '5511971918173' },
  { nome: 'Fábio', cpf: '21529074851', telefone: '5511989946586' },
];

const PROFESSORES_CADASTRO_DB = '728021ad4c58466db1dd5ab112ada252';

async function buscarProfessorPorCpf(cpfLimpo) {
  const r = await fetch('https://api.notion.com/v1/databases/' + PROFESSORES_CADASTRO_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ filter: { property: 'CPF', rich_text: { equals: cpfLimpo } }, page_size: 1 }),
  });
  const d = await r.json();
  const pagina = (d.results || [])[0];
  if (!pagina) return null;
  const p = pagina.properties;
  return {
    pageId: pagina.id,
    nome: p['Nome']?.title?.[0]?.plain_text || '',
    cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
    telefone: (p['Telefone']?.phone_number || '').replace(/\D/g, ''),
    endereco: p['Endereço']?.rich_text?.[0]?.plain_text || '',
    email: p['Email']?.email || '',
  };
}

// ===== PORTAL ADMIN — BUSCA POR NOME/CPF/TELEFONE =====
const INTEGRANTES_DB = 'e1047585-3dd2-4bda-9896-1a4caeeea284';

function normalizarBusca(texto) {
  return (texto || '').trim();
}

app.get('/portal-admin/busca', async (req, res) => {
  const tipo = req.query.tipo || '';
  const q = normalizarBusca(req.query.q);

  if (!q || q.length < 2) {
    return res.json({ ok: true, resultados: [] });
  }

  try {
    let resultados = [];

    if (tipo === 'aluna') {
      const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { or: [
            { property: 'Nome', title: { contains: q } },
            { property: 'CPF', rich_text: { contains: q } },
            { property: 'Contato', phone_number: { contains: q } },
          ]},
          page_size: 20,
        }),
      });
      const d = await r.json();
      resultados = (d.results || []).map(pagina => {
        const p = pagina.properties;
        return {
          pageId: pagina.id,
          nome: p['Nome']?.title?.[0]?.plain_text || '',
          cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
          telefone: p['Contato']?.phone_number || '',
          modalidade: p['Modalidade']?.select?.name || '',
          turma: p['Turma']?.select?.name || '',
          status: p['Status']?.select?.name || '',
        };
      });
    } else if (tipo === 'professor') {
      const r = await fetch('https://api.notion.com/v1/databases/' + PROFESSORES_CADASTRO_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { or: [
            { property: 'Nome', title: { contains: q } },
            { property: 'CPF', rich_text: { contains: q } },
            { property: 'Telefone', phone_number: { contains: q } },
          ]},
          page_size: 20,
        }),
      });
      const d = await r.json();
      resultados = (d.results || []).map(pagina => {
        const p = pagina.properties;
        return {
          pageId: pagina.id,
          nome: p['Nome']?.title?.[0]?.plain_text || '',
          cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
          telefone: p['Telefone']?.phone_number || '',
          email: p['Email']?.email || '',
        };
      });
    } else if (tipo === 'artista') {
      const r = await fetch('https://api.notion.com/v1/databases/' + INTEGRANTES_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { or: [
            { property: 'Nome', title: { contains: q } },
            { property: 'Nome Artístico', rich_text: { contains: q } },
            { property: 'CPF', rich_text: { contains: q } },
            { property: 'Telefone', phone_number: { contains: q } },
          ]},
          page_size: 20,
        }),
      });
      const d = await r.json();
      if (!r.ok) {
        console.error('[portal-admin/busca] erro Integrantes:', JSON.stringify(d));
        return res.status(500).json({ ok: false, erro: 'Erro ao consultar Integrantes. O token pode não ter acesso a esse banco — ver console.' });
      }
      resultados = (d.results || []).map(pagina => {
        const p = pagina.properties;
        return {
          pageId: pagina.id,
          nome: p['Nome']?.title?.[0]?.plain_text || '',
          nomeArtistico: p['Nome Artístico']?.rich_text?.[0]?.plain_text || '',
          cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
          telefone: p['Telefone']?.phone_number || '',
        };
      });
    } else {
      return res.status(400).json({ ok: false, erro: 'Tipo de busca inválido. Use aluna, professor ou artista.' });
    }

    res.json({ ok: true, resultados });
  } catch (err) {
    console.error('[portal-admin/busca] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar.' });
  }
});
// ===== FIM PORTAL ADMIN — BUSCA =====

async function buscarProfessorPorNome(nome) {
  const r = await fetch('https://api.notion.com/v1/databases/' + PROFESSORES_CADASTRO_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ filter: { property: 'Nome', title: { equals: nome } }, page_size: 1 }),
  });
  const d = await r.json();
  const pagina = (d.results || [])[0];
  if (!pagina) return null;
  return (pagina.properties['Telefone']?.phone_number || '').replace(/\D/g, '');
}

// ===== PORTAL ADMIN — DETALHE DE BUSCA (CARDS) =====

app.get('/portal-admin/busca/detalhe-aluna/:pageId', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + req.params.pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const pagina = await r.json();
    if (!r.ok) return res.status(404).json({ ok: false, erro: 'Aluna não encontrada.' });
    const p = pagina.properties;

    // Conta faltas nos últimos 30 dias, cruzando com Presenças 2026
    const trintaDiasAtras = new Date(Date.now() - 30 * 24 * 60 * 60000).toISOString();
    const rPresencas = await fetch('https://api.notion.com/v1/databases/8365a940-b386-401b-bedb-d26dfff2415e/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Status', select: { equals: 'Falta' } },
          { timestamp: 'created_time', created_time: { on_or_after: trintaDiasAtras } },
        ]},
        page_size: 100,
      }),
    });
    const dPresencas = await rPresencas.json();
    const faltas30dias = (dPresencas.results || []).filter(pp =>
      (pp.properties['Aluna']?.relation || []).some(rel => rel.id === pagina.id)
    ).length;

    res.json({
      ok: true,
      card: {
        nome: p['Nome']?.title?.[0]?.plain_text || '',
        cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
        telefone: p['Contato']?.phone_number || '',
        turma: p['Turma']?.select?.name || '',
        modalidade: p['Modalidade']?.select?.name || '',
        plano: p['Plano']?.select?.name || '',
        frequencia: p['Frequência']?.select?.name || '',
        vencimentoContrato: p['Vencimento do Contrato']?.date?.start || null,
        faltas30dias,
        linkContratoPdf: p['Link do Contrato PDF']?.url || '',
      },
    });
  } catch (err) {
    console.error('[busca/detalhe-aluna] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar detalhe.' });
  }
});

app.get('/portal-admin/busca/detalhe-professor/:pageId', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + req.params.pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const pagina = await r.json();
    if (!r.ok) return res.status(404).json({ ok: false, erro: 'Professor não encontrado.' });
    const p = pagina.properties;
    const nome = p['Nome']?.title?.[0]?.plain_text || '';

    // Busca o contrato de docência mais recente desse professor (match por nome)
    const rContrato = await fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROFESSORES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { property: 'Professor', rich_text: { equals: nome } },
        sorts: [{ timestamp: 'created_time', direction: 'descending' }],
        page_size: 1,
      }),
    });
    const dContrato = await rContrato.json();
    const contrato = (dContrato.results || [])[0];
    const cp = contrato ? contrato.properties : null;

    // Busca as turmas em Alunas onde esse professor está como titular (matrículas ativas)
    const rTurmas = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Professor', select: { equals: nome } },
          { property: 'Status', select: { equals: 'Ativa' } },
        ]},
        page_size: 100,
      }),
    });
    const dTurmas = await rTurmas.json();
    const turmasSet = new Set();
    (dTurmas.results || []).forEach(pagina2 => {
      const modalidade = pagina2.properties['Modalidade']?.select?.name || '';
      const turma = pagina2.properties['Turma']?.select?.name || '';
      if (turma) turmasSet.add(modalidade + ' — ' + turma);
    });

    res.json({
      ok: true,
      card: {
        nome,
        cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
        telefone: p['Telefone']?.phone_number || '',
        modalidades: (p['Modalidades']?.multi_select || []).map(m => m.name),
        turmas: [...turmasSet],
        trilha: cp ? (cp['Trilha']?.select?.name || '') : '',
        valorHoraAula: p['Valor Hora']?.number ?? (cp ? (cp['Valor Hora/Aula']?.number ?? null) : null),
        linkContratoPdf: cp ? (cp['Link do Contrato PDF']?.url || '') : '',
      },
    });
  } catch (err) {
    console.error('[busca/detalhe-professor] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar detalhe.' });
  }
});
// ===== FIM PORTAL ADMIN — DETALHE DE BUSCA (CARDS) =====

// ===== PORTAL ADMIN — PAINEL GERAL =====
const CONTRATOS_PROFESSORES_DB = 'f1b934a8100143019ac7f5f877c405f1';

app.get('/portal-admin/painel-geral', async (req, res) => {
  try {
    const hoje = new Date();
    const hojeISO = hoje.toISOString().slice(0, 10);
    const em30dias = new Date(hoje.getTime() - 30 * 24 * 60 * 60000).toISOString();
    const em14dias = new Date(hoje.getTime() + 14 * 24 * 60 * 60000).toISOString().slice(0, 10);

    // ---------- 1. Pendências ----------
    const [rAprov, rContratos, rSubs] = await Promise.all([
      fetch('https://api.notion.com/v1/databases/' + APROVACOES_ALUNAS_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ filter: { property: 'Status', select: { equals: 'Pendente' } }, page_size: 100 }),
      }),
      fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROFESSORES_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { or: [
            { property: 'Status', select: { equals: 'Aguardando Aceite' } },
            { property: 'Solicitação Pendente', select: { equals: 'Cancelamento' } },
            { property: 'Solicitação Pendente', select: { equals: 'Mudança' } },
          ]},
          page_size: 100,
        }),
      }),
      fetch('https://api.notion.com/v1/databases/' + SUBSTITUICOES_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { or: [
            { property: 'Status', select: { equals: 'Aguardando Confirmação' } },
            { property: 'Status', select: { equals: 'Sem Substituto - Fábio Notificado' } },
          ]},
          page_size: 100,
        }),
      }),
    ]);
    const dAprov = await rAprov.json();
    const dContratos = await rContratos.json();
    const dSubsPendentes = await rSubs.json();

    const pendencias = {
      aprovacoesAlunas: (dAprov.results || []).length,
      contratosProfessores: (dContratos.results || []).length,
      substituicoes: (dSubsPendentes.results || []).length,
    };
    pendencias.total = pendencias.aprovacoesAlunas + pendencias.contratosProfessores + pendencias.substituicoes;

    // ---------- 2. Aulas experimentais (últimos 30 dias) com presença ----------
    const rExp = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Status', select: { equals: 'Experimental' } },
          { timestamp: 'created_time', created_time: { on_or_after: em30dias } },
        ]},
        page_size: 100,
        sorts: [{ timestamp: 'created_time', direction: 'descending' }],
      }),
    });
    const dExp = await rExp.json();
    const experimentaisPaginas = dExp.results || [];

    const rPresencasExp = await fetch('https://api.notion.com/v1/databases/8365a940-b386-401b-bedb-d26dfff2415e/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ page_size: 100, sorts: [{ timestamp: 'created_time', direction: 'descending' }] }),
    });
    const dPresencasExp = await rPresencasExp.json();
    const presencasTodas = dPresencasExp.results || [];

    const experimentais = experimentaisPaginas.map(pagina => {
      const p = pagina.properties;
      const presencaMatch = presencasTodas.find(pp => (pp.properties['Aluna']?.relation || []).some(rel => rel.id === pagina.id));
      return {
        nome: p['Nome']?.title?.[0]?.plain_text || '',
        telefone: p['Contato']?.phone_number || '',
        modalidade: p['Modalidade']?.select?.name || '',
        turma: p['Turma']?.select?.name || '',
        criadoEm: pagina.created_time,
        presenca: presencaMatch ? (presencaMatch.properties['Status']?.select?.name || 'Aguardando') : 'Aguardando',
      };
    });

    // ---------- 3. Substituições nos próximos 14 dias ----------
    const rSubsProximas = await fetch('https://api.notion.com/v1/databases/' + SUBSTITUICOES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Data da Falta', date: { on_or_after: hojeISO } },
          { property: 'Data da Falta', date: { on_or_before: em14dias } },
        ]},
        page_size: 100,
        sorts: [{ property: 'Data da Falta', direction: 'ascending' }],
      }),
    });
    const dSubsProximas = await rSubsProximas.json();
    const substituicoesProximas = (dSubsProximas.results || []).map(pagina => {
      const p = pagina.properties;
      return {
        dataFalta: p['Data da Falta']?.date?.start || '',
        modalidade: p['Modalidade']?.select?.name || '',
        turma: p['Turma']?.rich_text?.[0]?.plain_text || '',
        professorTitular: p['Professor Titular']?.rich_text?.[0]?.plain_text || '',
        substituto: p['Substituto']?.rich_text?.[0]?.plain_text || '',
        status: p['Status']?.select?.name || '',
      };
    });

    // ---------- 4. Próximos feriados ----------
    const anoAtual = hoje.getFullYear();
    const feriadosSet = await getFeriadosDoAno(anoAtual);
    const feriadosProximoAno = await getFeriadosDoAno(anoAtual + 1);
    const todosFeriados = [...feriadosSet, ...feriadosProximoAno].sort();
    const proximosFeriados = todosFeriados.filter(f => f >= hojeISO).slice(0, 5);

    res.json({
      ok: true,
      pendencias,
      experimentais,
      substituicoesProximas,
      proximosFeriados,
    });
  } catch (err) {
    console.error('[portal-admin/painel-geral] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao montar o painel geral.' });
  }
});
// ===== FIM PORTAL ADMIN — PAINEL GERAL =====

// ===== AVALIAÇÃO CURSO DE ILUMINAÇÃO =====
const AVALIACAO_ILUMINACAO_DB = '28a27da340d1499293ba1bceccd5a736';

app.post('/avaliacao-luz/enviar', async (req, res) => {
  const b = req.body || {};

  function propSelect(valor) {
    return valor ? { select: { name: valor } } : { select: null };
  }
  function propText(valor) {
    return { rich_text: valor ? [{ text: { content: valor } }] : [] };
  }

  const propriedades = {
    'Título': { title: [{ text: { content: (b.nome || 'Respondente anônimo') + ' — ' + new Date().toLocaleDateString('pt-BR') } }] },
    'Deseja se Identificar': propSelect(b.desejaIdentificar),
    'Nome do Respondente': propText(b.nome),
    'Clareza dos Conceitos Tecnicos': propSelect(b.clareza),
    'Relevancia para a Pratica': propSelect(b.relevancia),
    'Profundidade do Conteudo': propSelect(b.profundidade),
    'Didatica do Instrutor': propSelect(b.didatica),
    'Qualidade do Material de Apoio': propSelect(b.material),
    'Tempo de Pratica com Equipamentos': propSelect(b.tempoPratica),
    'Qualidade dos Equipamentos': propSelect(b.qualidadeEquip),
    'Carga Horaria': propSelect(b.cargaHoraria),
    'Organizacao Geral': propSelect(b.organizacao),
    'Avaliacao do Espaco Fisico': propSelect(b.espacoFisico),
    'Avaliacao da Equipe de Atendimento': propSelect(b.equipeAtendimento),
    'Satisfacao Geral': propSelect(b.satisfacao),
    'Probabilidade de Recomendar (0-5)': { number: b.nps !== undefined && b.nps !== '' && b.nps !== null ? Number(b.nps) : null },
    'Nivel Previo em Iluminacao': propSelect(
      b.nivelPrevio === 'Intermediário' ? 'Intermediario' :
      b.nivelPrevio === 'Avançado' ? 'Avancado' :
      b.nivelPrevio
    ),
    'Faria Curso Complementar com Tomate': propSelect(b.tomate),
    'Faria Curso de Producao': propSelect(b.producao === 'Não' ? 'Nao' : b.producao),
    'Faria Curso de Tecnico de Som': propSelect(b.som === 'Não' ? 'Nao' : b.som),
    'O que mais gostou': propText(b.gostou),
    'O que melhoraria': propText(b.melhoraria),
  };

  try {
    const r = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: AVALIACAO_ILUMINACAO_DB }, properties: propriedades }),
    });
    if (!r.ok) {
      const t = await r.text();
      console.error('[avaliacao-luz/enviar] Notion recusou:', t);
      return res.status(500).json({ ok: false, erro: 'Erro ao salvar avaliação.' });
    }
    res.json({ ok: true });
  } catch (err) {
    console.error('[avaliacao-luz/enviar] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao salvar avaliação.' });
  }
});
// ===== FIM AVALIAÇÃO CURSO DE ILUMINAÇÃO =====

// ===== PORTAL ARTISTAS =====
const INTEGRANTES_DB_ARTISTA = 'e1047585-3dd2-4bda-9896-1a4caeeea284';

async function buscarIntegrantePorCpfArtista(cpfLimpo) {
  // O CPF em Integrantes pode estar salvo com ou sem pontuação (ex: 215.290.748-51),
  // então buscamos todos e comparamos só os números dos dois lados, em vez de
  // confiar num filtro exato do Notion que falharia com formatação diferente.
  const r = await fetch('https://api.notion.com/v1/databases/' + INTEGRANTES_DB_ARTISTA + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ page_size: 100 }),
  });
  const d = await r.json();
  const pagina = (d.results || []).find(pag => {
    const cpfArmazenado = (pag.properties['CPF']?.rich_text?.[0]?.plain_text || '').replace(/\D/g, '');
    return cpfArmazenado === cpfLimpo;
  });
  if (!pagina) return null;
  const p = pagina.properties;
  return {
    pageId: pagina.id,
    nome: p['Nome']?.title?.[0]?.plain_text || '',
    telefone: (p['Telefone']?.phone_number || '').replace(/\D/g, ''),
  };
}

app.post('/portal-artista/login/solicitar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  try {
    const integrante = await buscarIntegrantePorCpfArtista(cpfLimpo);
    if (!integrante) return res.json({ ok: false, erro: 'CPF não encontrado. Fale com o Fábio.' });
    if (!integrante.telefone) return res.json({ ok: false, erro: 'Telefone não cadastrado. Fale com o Fábio.' });
    await enviarOtp('artista_' + cpfLimpo, integrante.telefone, integrante.nome);
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-artista/login/solicitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-artista/login/verificar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  const verificacao = verificarOtp('artista_' + cpfLimpo, req.body.codigo);
  if (!verificacao.ok) return res.json(verificacao);
  try {
    const integrante = await buscarIntegrantePorCpfArtista(cpfLimpo);
    if (!integrante) return res.json({ ok: false, erro: 'Cadastro não encontrado.' });
    const dadosLogin = { nome: integrante.nome, cpf: cpfLimpo, integrantePageId: integrante.pageId };
    const sessionToken = criarSessao('artista_' + cpfLimpo, dadosLogin);
    res.json({ ok: true, ...dadosLogin, sessionToken });
  } catch (err) {
    console.error('[portal-artista/login/verificar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-artista/sessao/verificar', (req, res) => {
  const dados = verificarSessao(req.body.sessionToken);
  if (!dados) return res.json({ ok: false, erro: 'Sessão expirada.' });
  res.json({ ok: true, ...dados });
});

function calcularPrevisaoPagamento(emissaoNfISO, contratantes) {
  if (!emissaoNfISO) return { texto: 'Aguardando emissão da Nota Fiscal', data: null };
  const contratantesTexto = (contratantes || []).join(' ').toLowerCase();
  const prazoDias = (contratantesTexto.includes('sme') || contratantesTexto.includes('smc') || contratantesTexto.includes('prefeitura'))
    ? 30
    : 15;
  const dataNf = new Date(emissaoNfISO + 'T00:00:00');
  const previsao = new Date(dataNf.getTime() + prazoDias * 24 * 60 * 60000);
  const previsaoISO = previsao.toISOString().slice(0, 10);
  return { texto: previsaoISO.split('-').reverse().join('/'), data: previsaoISO, prazoDias };
}

app.get('/portal-artista/apresentacoes', async (req, res) => {
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });

  const ano = req.query.ano || String(new Date().getFullYear());
  if (ano !== '2026') {
    return res.json({ ok: true, apresentacoes: [], aviso: 'Ainda não verificamos se o banco de ' + ano + ' tem a mesma estrutura. Fale com o Fábio.' });
  }

  try {
    const idPessoa = dadosSessao.integrantePageId;
    const rApres = await fetch('https://api.notion.com/v1/databases/' + APRESENTACOES_2026_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { or: [
          { property: 'ELENCO', relation: { contains: idPessoa } },
          { property: 'Produção Liqui', relation: { contains: idPessoa } },
          { property: 'TÉCNICO DE SOM', relation: { contains: idPessoa } },
          { property: 'TÉCNICO DE LUZ', relation: { contains: idPessoa } },
        ]},
        page_size: 100,
        sorts: [{ property: 'Data da Apresentação', direction: 'descending' }],
      }),
    });
    const dApres = await rApres.json();
    const apresentacoesDoAno = (dApres.results || []).filter(pagina => {
      const dataStr = pagina.properties['Data da Apresentação']?.date?.start || '';
      return dataStr.slice(0, 4) === ano;
    });

    const resultado = [];
    for (const pagina of apresentacoesDoAno) {
      const p = pagina.properties;
      const localNome = p['LOCAL']?.title?.[0]?.plain_text || '';
      const dataApresentacao = p['Data da Apresentação']?.date?.start || '';
      const horario = p['Horário Apresentação']?.rich_text?.[0]?.plain_text || '';
      const fotosDrive = p['Fotos Drive']?.url || '';
      const propostaRel = p['Proposta']?.relation || [];

      // Determina o papel dessa pessoa NESSA apresentação específica
      const noElenco = (p['ELENCO']?.relation || []).some(r => r.id === idPessoa);
      const naProducao = (p['Produção Liqui']?.relation || []).some(r => r.id === idPessoa);
      const tecnicoSom = (p['TÉCNICO DE SOM']?.relation || []).some(r => r.id === idPessoa);
      const tecnicoLuz = (p['TÉCNICO DE LUZ']?.relation || []).some(r => r.id === idPessoa);
      let papel = 'Elenco';
      if (noElenco) papel = 'Elenco';
      else if (tecnicoSom || tecnicoLuz) papel = 'Técnico';
      else if (naProducao) papel = 'Produção';

      let cacheValor = null;
      let emissaoNf = null;
      let contratantes = [];
      let statusCache = '';

      if (propostaRel.length) {
        try {
          const rProposta = await fetch('https://api.notion.com/v1/pages/' + propostaRel[0].id, {
            headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
          });
          const propostaData = await rProposta.json();
          const pp = propostaData.properties || {};
          if (papel === 'Elenco') cacheValor = pp['Cachê Elenco']?.number ?? null;
          else if (papel === 'Produção') cacheValor = pp['Cachê Produção']?.number ?? null;
          else if (papel === 'Técnico') cacheValor = pp['Cachê Técnicos']?.number ?? null;
          emissaoNf = pp['Emissão da Nota Fiscal']?.date?.start || null;
          contratantes = (pp['Contratante']?.multi_select || []).map(o => o.name);
          statusCache = pp['Cachê']?.select?.name || '';
        } catch (e) {
          console.error('[portal-artista/apresentacoes] erro ao buscar proposta:', e.message);
        }
      }

      const previsao = calcularPrevisaoPagamento(emissaoNf, contratantes);

      resultado.push({
        local: localNome,
        data: dataApresentacao,
        horario,
        fotosDrive,
        papel,
        cacheElenco: cacheValor,
        statusCache: statusCache || 'Não informado',
        emissaoNf,
        contratantes,
        previsaoPagamentoTexto: previsao.texto,
      });
    }

    res.json({ ok: true, apresentacoes: resultado });
  } catch (err) {
    console.error('[portal-artista/apresentacoes] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar apresentações.' });
  }
});

app.get('/portal-artista/meus-dados', async (req, res) => {
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });

  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + dadosSessao.integrantePageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const pagina = await r.json();
    const p = pagina.properties;

    res.json({
      ok: true,
      dados: {
        nome: p['Nome']?.title?.[0]?.plain_text || '',
        nomeCompleto: p['Nome Completo']?.rich_text?.[0]?.plain_text || '',
        nomeSocial: p['Nome Social']?.rich_text?.[0]?.plain_text || '',
        nomeArtistico: p['Nome Artístico']?.rich_text?.[0]?.plain_text || '',
        cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
        rg: p['RG']?.rich_text?.[0]?.plain_text || '',
        dataNascimento: p['Data de Nascimento']?.date?.start || '',
        telefone: p['Telefone']?.phone_number || '',
        email: p['E-mail']?.email || '',
        endereco: p['Endereço']?.rich_text?.[0]?.plain_text || '',
        instagram: p['Instagram']?.rich_text?.[0]?.plain_text || '',
        pix: p['PIX']?.rich_text?.[0]?.plain_text || '',
        dadosBancarios: p['Dados Bancários']?.rich_text?.[0]?.plain_text || '',
        pis: p['PIS']?.number ?? '',
        altura: p['Altura']?.number ?? '',
        sapato: p['Sapato']?.number ?? '',
        camisa: p['Camisa']?.rich_text?.[0]?.plain_text || '',
        calca: p['Calça']?.rich_text?.[0]?.plain_text || '',
        camiseta: p['Camiseta']?.rich_text?.[0]?.plain_text || '',
        etnia: p['Etnia']?.rich_text?.[0]?.plain_text || '',
        genero: p['Gênero']?.rich_text?.[0]?.plain_text || '',
        pronomes: p['Pronomes']?.rich_text?.[0]?.plain_text || '',
        orientacaoSexual: p['Orientação Sexual']?.rich_text?.[0]?.plain_text || '',
        pcd: p['PCD']?.rich_text?.[0]?.plain_text || '',
        alergias: p['Alergias']?.rich_text?.[0]?.plain_text || '',
        restricaoAlimentar: p['Restrição Alimentar']?.rich_text?.[0]?.plain_text || '',
        contatoEmergencia: p['Contato de Emergência']?.rich_text?.[0]?.plain_text || '',
        drtOmb: p['DRT / OMB']?.rich_text?.[0]?.plain_text || '',
      },
    });
  } catch (err) {
    console.error('[portal-artista/meus-dados] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar dados.' });
  }
});

app.post('/portal-artista/atualizar-dados', async (req, res) => {
  const dadosSessao = verificarSessao(req.body.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });

  const b = req.body || {};
  function txt(v) { return { rich_text: v ? [{ text: { content: String(v) } }] : [] }; }
  function num(v) { return { number: v !== undefined && v !== null && v !== '' ? Number(v) : null }; }

  const propriedades = {
    'Nome Completo': txt(b.nomeCompleto),
    'Nome Social': txt(b.nomeSocial),
    'Nome Artístico': txt(b.nomeArtistico),
    'RG': txt(b.rg),
    'Telefone': { phone_number: b.telefone || null },
    'E-mail': { email: b.email || null },
    'Endereço': txt(b.endereco),
    'Instagram': txt(b.instagram),
    'PIX': txt(b.pix),
    'Dados Bancários': txt(b.dadosBancarios),
    'PIS': num(b.pis),
    'Altura': num(b.altura),
    'Sapato': num(b.sapato),
    'Camisa': txt(b.camisa),
    'Calça': txt(b.calca),
    'Camiseta': txt(b.camiseta),
    'Etnia': txt(b.etnia),
    'Gênero': txt(b.genero),
    'Pronomes': txt(b.pronomes),
    'Orientação Sexual': txt(b.orientacaoSexual),
    'PCD': txt(b.pcd),
    'Alergias': txt(b.alergias),
    'Restrição Alimentar': txt(b.restricaoAlimentar),
    'Contato de Emergência': txt(b.contatoEmergencia),
    'DRT / OMB': txt(b.drtOmb),
  };
  if (b.dataNascimento) propriedades['Data de Nascimento'] = { date: { start: b.dataNascimento } };

  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + dadosSessao.integrantePageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: propriedades }),
    });
    if (!r.ok) {
      const t = await r.text();
      console.error('[portal-artista/atualizar-dados] Notion recusou:', t);
      return res.status(500).json({ ok: false, erro: 'Erro ao salvar dados.' });
    }
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-artista/atualizar-dados] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao salvar dados.' });
  }
});
// ===== FIM PORTAL ARTISTAS =====
// ===== PORTAL ARTISTAS — EXTRAS =====

// 1. Contrato de docência (se o artista também for professor cadastrado)
app.get('/portal-artista/contrato-professor', async (req, res) => {
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });

  try {
    const rProf = await fetch('https://api.notion.com/v1/databases/' + PROFESSORES_CADASTRO_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'CPF', rich_text: { equals: dadosSessao.cpf } }, page_size: 1 }),
    });
    const dProf = await rProf.json();
    const profPagina = (dProf.results || [])[0];
    if (!profPagina) return res.json({ ok: true, ehProfessor: false });

    const nomeProf = profPagina.properties['Nome']?.title?.[0]?.plain_text || '';

    const rContrato = await fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROFESSORES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { property: 'Professor', rich_text: { equals: nomeProf } },
        sorts: [{ timestamp: 'created_time', direction: 'descending' }],
        page_size: 1,
      }),
    });
    const dContrato = await rContrato.json();
    const contratoPagina = (dContrato.results || [])[0];
    if (!contratoPagina) return res.json({ ok: true, ehProfessor: true, contrato: null });

    const cp = contratoPagina.properties;
    res.json({
      ok: true,
      ehProfessor: true,
      contrato: {
        trilha: cp['Trilha']?.select?.name || '',
        status: cp['Status']?.select?.name || '',
        linkContratoPdf: cp['Link do Contrato PDF']?.url || '',
      },
    });
  } catch (err) {
    console.error('[portal-artista/contrato-professor] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar contrato.' });
  }
});

// 2. Meu Acervo — itens físicos (figurinos/props/equipamentos) vinculados ao artista
app.get('/portal-artista/meu-acervo', async (req, res) => {
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });

  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ACERVO_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { property: 'Integrantes', relation: { contains: dadosSessao.integrantePageId } },
        page_size: 100,
      }),
    });
    const d = await r.json();
    console.log('[portal-artista/meu-acervo] DEBUG idPessoa=' + dadosSessao.integrantePageId + ' | r.ok=' + r.ok + ' | total resultados=' + (d.results || []).length + ' | resposta bruta=' + JSON.stringify(d).slice(0, 500));
    const itens = (d.results || []).map(pagina => {
      const p = pagina.properties;
      return {
        nome: p['Nome']?.title?.[0]?.plain_text || '',
        tipo: (p['Tipo']?.multi_select || []).map(o => o.name),
        status: p['STATUS']?.status?.name || '',
        medidas: p['Medidas']?.rich_text?.[0]?.plain_text || '',
        numeracao: p['Numeração']?.rich_text?.[0]?.plain_text || '',
        validade: p['Validade']?.rich_text?.[0]?.plain_text || '',
      };
    });
    res.json({ ok: true, itens });
  } catch (err) {
    console.error('[portal-artista/meu-acervo] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar acervo.' });
  }
});

// 3. Notificação automática de pagamento (webhook disparado quando Cachê vira PAGO)
app.post('/webhook-cache-pago', async (req, res) => {
  res.status(200).json({ ok: true });

  try {
    const body = req.body || {};
    const propostaPageId = (body.data && body.data.id) || body.pageId || body.page_id || null;
    if (!propostaPageId) { console.error('[webhook-cache-pago] sem page id.'); return; }

    const rProposta = await fetch('https://api.notion.com/v1/pages/' + propostaPageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const proposta = await rProposta.json();
    const pp = proposta.properties;

    const statusCache = pp['Cachê']?.select?.name;
    const jaNotificado = pp['Cache Notificado']?.checkbox || false;
    if (statusCache !== 'PAGO' || jaNotificado) {
      console.log('[webhook-cache-pago] ignorando (status=' + statusCache + ', jaNotificado=' + jaNotificado + ')');
      return;
    }

    // Acha as apresentações vinculadas a essa proposta, pra saber o elenco
    const rApres = await fetch('https://api.notion.com/v1/databases/' + APRESENTACOES_2026_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Proposta', relation: { contains: propostaPageId } }, page_size: 20 }),
    });
    const dApres = await rApres.json();
    const localNome = (dApres.results || [])[0]?.properties?.['LOCAL']?.title?.[0]?.plain_text || 'sua apresentação';

    const elencoIds = new Set();
    (dApres.results || []).forEach(pagina => {
      (pagina.properties['ELENCO']?.relation || []).forEach(rel => elencoIds.add(rel.id));
      (pagina.properties['Produção Liqui']?.relation || []).forEach(rel => elencoIds.add(rel.id));
      (pagina.properties['TÉCNICO DE SOM']?.relation || []).forEach(rel => elencoIds.add(rel.id));
      (pagina.properties['TÉCNICO DE LUZ']?.relation || []).forEach(rel => elencoIds.add(rel.id));
    });

    const msg = '💰 Boa notícia! Seu cachê da apresentação em *' + localNome + '* foi *pago*! ' +
      'Confere o valor certinho no Portal Equipe. Qualquer dúvida, é só chamar. 🎭';

    for (const integranteId of elencoIds) {
      try {
        const rInt = await fetch('https://api.notion.com/v1/pages/' + integranteId, {
          headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
        });
        const integrante = await rInt.json();
        const telefone = (integrante.properties['Telefone']?.phone_number || '').replace(/\D/g, '');
        if (telefone) await enviarWhatsAppComHorarioComercial('55' + telefone, msg);
      } catch (e) {
        console.error('[webhook-cache-pago] erro ao notificar integrante ' + integranteId + ':', e.message);
      }
    }

    await fetch('https://api.notion.com/v1/pages/' + propostaPageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: { 'Cache Notificado': { checkbox: true } } }),
    });

    console.log('[webhook-cache-pago] notificados ' + elencoIds.size + ' integrante(s) sobre pagamento.');
  } catch (err) {
    console.error('[webhook-cache-pago] erro:', err.message);
  }
});
// ===== FIM PORTAL ARTISTAS — EXTRAS =====






// ===== LISTA DE INTERESSE — CURSOS SEM TURMA ABERTA =====
const LISTA_INTERESSE_DB = 'd0ffcc65-5f9b-4206-8312-6d288fd56ad9';

app.post('/interesse-curso', async (req, res) => {
  const { curso, nome, whatsapp, mensagem, melhorDiaHorario } = req.body;
  if (!curso || !nome || !whatsapp) {
    return res.status(400).json({ ok: false, erro: 'Preencha nome e WhatsApp.' });
  }
  try {
    await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: LISTA_INTERESSE_DB },
        properties: {
          'Título': { title: [{ text: { content: nome + ' — ' + curso } }] },
          'Curso': { select: { name: curso } },
          'Nome': { rich_text: [{ text: { content: nome } }] },
          'WhatsApp': { phone_number: whatsapp },
          'Melhor Dia Horario': { rich_text: [{ text: { content: melhorDiaHorario || '' } }] },
          'Mensagem': { rich_text: [{ text: { content: mensagem || '' } }] },
          'Contatado': { checkbox: false },
        },
      }),
    });

    const msgFabio = '📋 *Novo interesse — ' + curso + '*\n\nNome: ' + nome + '\nWhatsApp: ' + whatsapp + (melhorDiaHorario ? ('\nMelhor dia/horário: ' + melhorDiaHorario) : '') + (mensagem ? ('\nMensagem: ' + mensagem) : '');
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgFabio); } catch (e) {}

    res.json({ ok: true });
  } catch (err) {
    console.error('[interesse-curso] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao registrar interesse.' });
  }
});
// ===== FIM LISTA DE INTERESSE =====

// ===== TERMO DE RESIDÊNCIA ARTÍSTICA 2026 =====
const RESIDENCIA_INSCRICOES_DB = 'a514d3c1-1a83-40ff-bc73-fa30582f6296';

function mapaDatasPretendidas(valor) {
  return {
    terca: valor === 'Terça',
    quinta: valor === 'Quinta',
    ambas: valor === 'Ambas',
  };
}

function montarTextoTermoResidencia({ nomeResponsavel, rgResponsavel, cpfResponsavel, enderecoResponsavel, nomeProjeto, datasPretendidas, detalhamentoAtividade }) {
  const datas = mapaDatasPretendidas(datasPretendidas);
  const checkboxDatas =
    '(' + (datas.terca ? 'X' : ' ') + ') Terças-feiras, das 10h às 13h\n' +
    '(' + (datas.quinta ? 'X' : ' ') + ') Quintas-feiras, das 10h às 13h\n' +
    '(' + (datas.ambas ? 'X' : ' ') + ') Terças e quintas-feiras, das 10h às 13h';

  return 'TERMO DE RESIDÊNCIA ARTÍSTICA\nEspaço Liquidificador — 2º semestre de 2026\n\n' +
    'ESPAÇO/CONTRATANTE: ESPAÇO LIQUIDIFICADOR, nome fantasia de LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS, pessoa jurídica inscrita no CNPJ sob nº 28.398.119/0001-83, com sede na Rua Dr. Carvalho de Mendonça, 67, Campos Elíseos, São Paulo/SP, neste ato representada por Cristiane Socci Leonel, doravante denominado simplesmente "ESPAÇO".\n\n' +
    'RESPONSÁVEL PELO PROJETO/RESIDENTE: ' + nomeResponsavel + ', RG ' + rgResponsavel + ', CPF ' + cpfResponsavel + ', residente e domiciliado(a) em ' + enderecoResponsavel + ', atuando em nome próprio e representando o coletivo/projeto "' + nomeProjeto + '", composto pelos integrantes listados no Anexo II, doravante denominado simplesmente "RESIDENTE".\n\n' +
    'As partes acima identificadas celebram o presente Termo de Residência Artística, decorrente da seleção do projeto no Chamamento de Residência Artística — Espaço Liquidificador 2º semestre de 2026 ("Chamamento"), que mediante as cláusulas a seguir reciprocamente outorgam e aceitam.\n\n' +
    'CLÁUSULA PRIMEIRA — DO OBJETO\n\n' +
    '1.1. O presente Termo tem por objeto formalizar a cessão de uso não onerosa, pelo ESPAÇO ao RESIDENTE, do espaço físico descrito no item 1.8 do Chamamento, para fins exclusivos de pesquisa, elaboração e/ou ensaio do projeto "' + nomeProjeto + '", nos termos e condições do Chamamento, que integra este instrumento como Anexo III, independentemente de transcrição.\n\n' +
    '1.2. Em caso de conflito entre este Termo e o Chamamento, prevalecerão as disposições deste Termo quanto às condições específicas da residência do RESIDENTE.\n\n' +
    'CLÁUSULA SEGUNDA — DA VIGÊNCIA E DAS DATAS DE USO\n\n' +
    '2.1. Este Termo vigora de 11 de agosto a 17 de dezembro de 2026.\n\n' +
    '2.2. O uso do espaço ocorrerá nas seguintes datas e horários, conforme resultado da seleção:\n\n' + checkboxDatas + '\n\n' +
    '2.3. Eventuais alterações pontuais de data ou horário observarão o item 1.5 do Chamamento, buscando-se sempre o comum acordo e a reposição, sem prejuízo à continuidade do projeto.\n\n' +
    '2.4. O uso do espaço fora das datas e horários acima, ou a extensão do período de residência, dependerá de acordo específico entre as partes, com cobrança de valor à parte, conforme disponibilidade de agenda do ESPAÇO (item 1.4 do Chamamento).\n\n' +
    'CLÁUSULA TERCEIRA — DA NATUREZA DA RELAÇÃO\n\n' +
    '3.1. O presente Termo formaliza cessão de uso de espaço para fins culturais, no contexto de residência artística selecionada por chamamento público, não configurando relação de trabalho, prestação de serviços onerosa, sociedade, associação ou vínculo empregatício de qualquer natureza entre o ESPAÇO e o RESIDENTE ou seus integrantes.\n\n' +
    '3.2. Não há entre as partes relação de subordinação, pessoalidade, habitualidade remunerada ou exclusividade. O RESIDENTE exerce as atividades do projeto com plena autonomia artística, técnica e organizacional, sem qualquer supervisão, direção ou fiscalização de conteúdo por parte do ESPAÇO, cabendo a este apenas a gestão do uso do espaço físico e a verificação do cumprimento das contrapartidas pactuadas.\n\n' +
    '3.3. O uso gratuito do espaço não constitui contraprestação por trabalho, e as contrapartidas previstas na Cláusula Quinta não têm natureza de prestação de serviço oneroso, mas de condição de reciprocidade cultural e comunitária inerente ao programa de residência.\n\n' +
    'CLÁUSULA QUARTA — DAS CONDIÇÕES DE USO DO ESPAÇO\n\n' +
    '4.1. O RESIDENTE terá acesso à sala de ensaio, banheiros, cozinha e à área de armazenamento de materiais indicada pelo ESPAÇO, conforme item 1.6 do Chamamento, sendo vedado o acesso ao escritório/mezanino e às demais dependências do acervo (item 1.7 do Chamamento).\n\n' +
    '4.2. O RESIDENTE compromete-se a utilizar o espaço e os equipamentos disponibilizados (item 1.8 do Chamamento) de forma cuidadosa e conforme sua finalidade, respondendo por danos nos termos da Cláusula Sétima.\n\n' +
    '4.3. O uso de aparelhos aéreos e demais equipamentos de risco rege-se pelo Anexo I — Termo de Responsabilidade, Uso e Riscos, assinado em conjunto com este Termo e dele parte integrante.\n\n' +
    'CLÁUSULA QUINTA — DAS CONTRAPARTIDAS\n\n' +
    '5.1. Contrapartidas artísticas e de divulgação. O RESIDENTE compromete-se a:\n' +
    'a) registrar a residência e enviar ao ESPAÇO relatório final das atividades realizadas, acompanhado de fotos e vídeos, até 30 (trinta) dias após o encerramento da vigência;\n' +
    'b) realizar 1 (uma) atividade de formação aberta (oficina, roda de conversa, palestra, workshop ou similar), na seguinte carga horária, dia(s) e horário(s): ' + detalhamentoAtividade + ', conforme proposta apresentada na inscrição, podendo ser ajustada de comum acordo quanto à data;\n' +
    'c) inserir a logomarca do Espaço Liquidificador em todo material de divulgação do projeto, na condição de "apoio";\n' +
    'd) mencionar/marcar os perfis oficiais do Espaço Liquidificador nas publicações em redes sociais relativas ao registro da residência.\n\n' +
    '5.2. Contrapartida de cuidado coletivo do espaço. O RESIDENTE assume a co-responsabilidade pelo cuidado do ambiente compartilhado, nos exatos termos do item 4.2 do Chamamento, aqui incorporado por referência, destacando-se que:\n' +
    'a) a contribuição corresponde a aproximadamente 2 (duas) horas semanais totais, independentemente do número de datas utilizadas;\n' +
    'b) não inclui manutenção técnica, estrutural ou de equipamentos de risco, reparos, nem tarefas que substituam pessoal contratado do ESPAÇO;\n' +
    'c) tem caráter colaborativo, não configurando relação de trabalho, subordinação, pessoalidade ou vínculo de qualquer natureza, em reforço ao disposto na Cláusula Terceira.\n\n' +
    '5.3. O descumprimento reiterado e não justificado das contrapartidas desta cláusula poderá ensejar as consequências previstas na Cláusula Oitava.\n\n' +
    'CLÁUSULA SEXTA — DA AUTORIZAÇÃO DE USO DE IMAGEM E DIREITOS DE DIVULGAÇÃO\n\n' +
    '6.1. O RESIDENTE autoriza, a título gratuito e não exclusivo, o uso pelo ESPAÇO de fotos, vídeos e demais registros da residência — inclusive os enviados no relatório final (item 5.1, "a") e os produzidos pelo próprio ESPAÇO durante o período de uso do espaço — para fins de divulgação institucional do Espaço Liquidificador e da Cia. do Liquidificador em suas redes sociais, site e materiais de prestação de contas, com a devida atribuição de crédito ao projeto e a seus integrantes.\n\n' +
    '6.2. Reciprocamente, o ESPAÇO autoriza o RESIDENTE a utilizar, para fins de divulgação do próprio projeto, imagens do espaço físico captadas durante o período de residência, vedada a associação a conteúdo que prejudique a imagem do ESPAÇO.\n\n' +
    '6.3. Esta autorização não transfere direitos autorais sobre a obra artística em desenvolvimento, que permanecem integralmente com o RESIDENTE e seus integrantes.\n\n' +
    'CLÁUSULA SÉTIMA — DA VISTORIA E DOS DANOS\n\n' +
    '7.1. Será realizada vistoria de entrada, no início da residência, e vistoria de saída, ao término da vigência ou da última data de uso, registrando o estado da sala e dos equipamentos, conforme item 7.3 do Chamamento.\n\n' +
    '7.2. Danos causados ao espaço ou a seus equipamentos, apurados por comparação entre as vistorias de entrada e saída, serão ressarcidos pelo RESIDENTE, mediante notificação e orçamento apresentado pelo ESPAÇO, no prazo de 15 (quinze) dias corridos.\n\n' +
    'CLÁUSULA OITAVA — DA RESCISÃO E DAS PENALIDADES\n\n' +
    '8.1. Constituem hipóteses de rescisão deste Termo, mediante notificação por escrito:\n' +
    'a) descumprimento do item 3.4 do Chamamento (realização de apresentações, oficinas ou cursos como objeto da residência), hipótese em que será também aplicada multa de R$ 500,00 (quinhentos reais), conforme item 8.2 do Chamamento;\n' +
    'b) descumprimento reiterado e não sanado, após notificação e prazo de 10 (dez) dias corridos para regularização, das contrapartidas previstas na Cláusula Quinta;\n' +
    'c) uso do espaço em desacordo com sua finalidade, com risco à integridade física de terceiros ou ao patrimônio do ESPAÇO;\n' +
    'd) acordo entre as partes ou desistência formal do RESIDENTE, mediante comunicação com antecedência mínima de 15 (quinze) dias.\n\n' +
    '8.2. A rescisão não gera direito a indenização entre as partes, ressalvado o ressarcimento de danos eventualmente apurados nos termos da Cláusula Sétima.\n\n' +
    'CLÁUSULA NONA — DA PROTEÇÃO DE DADOS PESSOAIS (LGPD)\n\n' +
    '9.1. O tratamento de dados pessoais do RESIDENTE e de seus integrantes no âmbito da residência observa a Lei nº 13.709/2018 (LGPD) e segue as finalidades e condições descritas na Seção 6 do Chamamento, aqui incorporadas por referência.\n\n' +
    'CLÁUSULA DÉCIMA — DAS DISPOSIÇÕES GERAIS\n\n' +
    '10.1. Este Termo, seu Anexo I (Termo de Responsabilidade, Uso e Riscos) e o Chamamento constituem o inteiro acordo entre as partes quanto à residência artística ora formalizada.\n\n' +
    '10.2. O ESPAÇO é identificado publicamente pelo nome fantasia "Espaço Liquidificador", sendo LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS a pessoa jurídica signatária e responsável pelas obrigações aqui assumidas.\n\n' +
    '10.3. Este Termo não poderá ser cedido ou transferido pelo RESIDENTE a terceiros sem prévia anuência por escrito do ESPAÇO.\n\n' +
    'CLÁUSULA DÉCIMA PRIMEIRA — DO FORO\n\n' +
    '11.1. Fica eleito o foro da Comarca de São Paulo/SP para dirimir quaisquer controvérsias oriundas deste Termo, com renúncia a qualquer outro, por mais privilegiado que seja.';
}

async function buscarProjetoResidenciaPorCpf(cpfLimpo) {
  const r = await fetch('https://api.notion.com/v1/databases/' + RESIDENCIA_INSCRICOES_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ page_size: 100 }),
  });
  const d = await r.json();
  const pagina = (d.results || []).find(pag => {
    const cpfArmazenado = (pag.properties['CPF']?.rich_text?.[0]?.plain_text || '').replace(/\D/g, '');
    return cpfArmazenado === cpfLimpo;
  });
  if (!pagina) return null;
  const p = pagina.properties;
  return {
    pageId: pagina.id,
    nome: p['Nome Civil']?.rich_text?.[0]?.plain_text || '',
    telefone: (p['Telefone']?.phone_number || '').replace(/\D/g, ''),
    status: p['Status']?.select?.name || '',
  };
}

app.post('/residencia-termo/login/solicitar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  try {
    const projeto = await buscarProjetoResidenciaPorCpf(cpfLimpo);
    if (!projeto) return res.json({ ok: false, erro: 'CPF não encontrado entre os responsáveis. Fale com o Fábio.' });
    if (projeto.status !== 'Selecionada') return res.json({ ok: false, erro: 'Esse CPF não corresponde a um projeto selecionado.' });
    if (!projeto.telefone) return res.json({ ok: false, erro: 'Telefone não cadastrado. Fale com o Fábio.' });
    await enviarOtp('residencia_' + cpfLimpo, projeto.telefone, projeto.nome);
    res.json({ ok: true });
  } catch (err) {
    console.error('[residencia-termo/login/solicitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/residencia-termo/login/verificar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  const verificacao = verificarOtp('residencia_' + cpfLimpo, req.body.codigo);
  if (!verificacao.ok) return res.json(verificacao);
  try {
    const projeto = await buscarProjetoResidenciaPorCpf(cpfLimpo);
    if (!projeto) return res.json({ ok: false, erro: 'Projeto não encontrado.' });
    const dadosLogin = { nome: projeto.nome, cpf: cpfLimpo, projetoPageId: projeto.pageId };
    const sessionToken = criarSessao('residencia_' + cpfLimpo, dadosLogin);
    res.json({ ok: true, ...dadosLogin, sessionToken });
  } catch (err) {
    console.error('[residencia-termo/login/verificar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/residencia-termo/sessao/verificar', (req, res) => {
  const dados = verificarSessao(req.body.sessionToken);
  if (!dados) return res.json({ ok: false, erro: 'Sessão expirada.' });
  res.json({ ok: true, ...dados });
});

app.get('/residencia-termo/dados', async (req, res) => {
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  const pageId = dadosSessao.projetoPageId;
  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    if (!r.ok) return res.json({ ok: false, erro: 'Projeto não encontrado.' });
    const pagina = await r.json();
    const p = pagina.properties;

    res.json({
      ok: true,
      nomeCivil: p['Nome Civil']?.rich_text?.[0]?.plain_text || '',
      rg: p['RG']?.rich_text?.[0]?.plain_text || '',
      cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
      nomeProjeto: p['Nome do Projeto']?.rich_text?.[0]?.plain_text || '',
      nomeColetivo: p['Nome do Coletivo']?.rich_text?.[0]?.plain_text || '',
      datasPretendidas: p['Datas Pretendidas']?.select?.name || '',
      propostaContrapartida: p['Proposta de Contrapartida']?.rich_text?.[0]?.plain_text || '',
      termoJaAceito: p['Termo Residencia Assinado']?.checkbox || false,
      linkTermoPdf: p['Link Termo PDF']?.url || '',
    });
  } catch (err) {
    console.error('[residencia] erro ao buscar projeto:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar projeto.' });
  }
});

app.post('/residencia-termo/aceitar', async (req, res) => {
  const dadosSessao = verificarSessao(req.body.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  const pageId = dadosSessao.projetoPageId;
  const { enderecoResponsavel, detalhamentoAtividade, assinaturaDigitada, aceite, dispositivo } = req.body;
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

  if (!enderecoResponsavel || !assinaturaDigitada || !aceite) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  const detalhamentoFinal = (detalhamentoAtividade || '').trim() || 'a ser definido entre as partes ao longo do período de residência, com antecedência mínima de 15 dias da data da atividade';

  try {
    const rPagina = await fetch('https://api.notion.com/v1/pages/' + pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    if (!rPagina.ok) return res.json({ ok: false, erro: 'Projeto não encontrado.' });
    const pagina = await rPagina.json();
    const p = pagina.properties;

    const nomeResponsavel = p['Nome Civil']?.rich_text?.[0]?.plain_text || '';
    const rgResponsavel = p['RG']?.rich_text?.[0]?.plain_text || '';
    const cpfResponsavel = p['CPF']?.rich_text?.[0]?.plain_text || '';
    const nomeProjeto = p['Nome do Projeto']?.rich_text?.[0]?.plain_text || '';
    const datasPretendidas = p['Datas Pretendidas']?.select?.name || '';
    const telefone = p['Telefone']?.phone_number || '';
    const email = p['Email']?.email || '';

    const textoTermo = montarTextoTermoResidencia({
      nomeResponsavel, rgResponsavel, cpfResponsavel, enderecoResponsavel,
      nomeProjeto, datasPretendidas, detalhamentoAtividade: detalhamentoFinal,
    });

    const dataHoraISO = new Date().toISOString();
    const pdfBuffer = await gerarPdfContrato(textoTermo, { nome: nomeResponsavel, dataHoraISO, ip, dispositivo, assinaturaDigitada });

    const msToken = await getMicrosoftToken();
    const nomePasta = 'Residencia-' + slugify(nomeProjeto);
    const folderId = await criarOuObterSubpasta(msToken, nomePasta);
    const pdfUploaded = await uploadBufferOneDrive(msToken, folderId, pdfBuffer, 'termo-residencia-assinado-' + slugify(nomeProjeto) + '.pdf');
    const linkPdf = await criarLinkCompartilhamento(msToken, pdfUploaded.id);

    await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Endereco Responsavel': { rich_text: [{ text: { content: enderecoResponsavel } }] },
          'Detalhamento Atividade Formativa': { rich_text: [{ text: { content: detalhamentoFinal } }] },
          'Termo Residencia Assinado': { checkbox: true },
          'Link Termo PDF': { url: linkPdf },
        },
      }),
    });

    const primeiroNome = nomeResponsavel.split(' ')[0];
    const numLimpo = telefone.replace(/\D/g, '');
    const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
    if (numBr) {
      try { await enviarWhatsApp(numBr, 'Prontinho, ' + primeiroNome + '! ✅\n\nO Termo de Residência Artística do projeto "' + nomeProjeto + '" foi assinado com sucesso. Aqui está o PDF:\n' + linkPdf); } catch(e) {}
    }
    try { await enviarWhatsApp(WHATSAPP_FABIO, '📜 Termo de Residência assinado — "' + nomeProjeto + '" (' + nomeResponsavel + ')\nPDF: ' + linkPdf); } catch(e) {}

    res.json({ ok: true, linkPdf });
  } catch (err) {
    console.error('[residencia] erro ao aceitar:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});
// ===== FIM TERMO DE RESIDÊNCIA ARTÍSTICA =====





// ===== CALCULADORA DE DESLOCAMENTO — DISTÂNCIA E PEDÁGIO VIA GOOGLE ROUTES API =====
const GOOGLE_ROUTES_API_KEY = process.env.GOOGLE_ROUTES_API_KEY;
const ENDERECO_ORIGEM_ESPACO = 'Rua Dr. Carvalho de Mendonça, 67, Campos Elíseos, São Paulo, SP, Brasil';

app.post('/calcular-distancia-deslocamento', async (req, res) => {
  const { enderecoDestino } = req.body;
  if (!enderecoDestino) return res.status(400).json({ ok: false, erro: 'Informe o endereço de destino.' });
  if (!GOOGLE_ROUTES_API_KEY) return res.status(500).json({ ok: false, erro: 'Chave da Google Routes API não configurada no servidor.' });

  try {
    const body = {
      origin: { address: ENDERECO_ORIGEM_ESPACO },
      destination: { address: enderecoDestino },
      travelMode: 'DRIVE',
      routingPreference: 'TRAFFIC_AWARE',
      extraComputations: ['TOLLS'],
      routeModifiers: {
        vehicleInfo: { emissionType: 'GASOLINE' },
      },
      units: 'METRIC',
    };

    const r = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': GOOGLE_ROUTES_API_KEY,
        'X-Goog-FieldMask': 'routes.distanceMeters,routes.duration,routes.travelAdvisory.tollInfo,routes.legs.steps.distanceMeters,routes.legs.steps.staticDuration',
      },
      body: JSON.stringify(body),
    });

    const d = await r.json();
    if (!r.ok) {
      console.error('[calcular-distancia-deslocamento] erro Google:', JSON.stringify(d));
      return res.status(500).json({ ok: false, erro: d.error?.message || 'Erro ao consultar rota.' });
    }

    const rota = (d.routes || [])[0];
    if (!rota) return res.json({ ok: false, erro: 'Não foi possível calcular a rota até esse endereço.' });

    const distanciaIdaKm = rota.distanceMeters / 1000;
    const distanciaTotalKm = distanciaIdaKm * 2; // ida + volta

    const tollInfo = rota.travelAdvisory?.tollInfo;
    let pedagioIda = 0;
    if (tollInfo?.estimatedPrice?.length) {
      // Soma todas as moedas retornadas (normalmente só BRL)
      const brl = tollInfo.estimatedPrice.find(p => p.currencyCode === 'BRL');
      if (brl) pedagioIda = Number(brl.units || 0) + Number(brl.nanos || 0) / 1e9;
    }
    const pedagioTotal = pedagioIda * 2; // ida + volta

    // Estima %cidade vs %estrada olhando a velocidade de cada trecho (step) da rota.
    // Acima de 60km/h no trecho, consideramos rodovia; abaixo, consideramos cidade.
    const LIMIAR_VELOCIDADE_RODOVIA_KMH = 60;
    let metrosCidade = 0;
    let metrosEstrada = 0;
    const steps = (rota.legs || []).flatMap(leg => leg.steps || []);
    steps.forEach(step => {
      const metros = step.distanceMeters || 0;
      const duracaoSeg = parseInt(step.staticDuration) || 0;
      if (metros === 0) return;
      const velocidadeKmh = duracaoSeg > 0 ? (metros / 1000) / (duracaoSeg / 3600) : 0;
      if (velocidadeKmh >= LIMIAR_VELOCIDADE_RODOVIA_KMH) metrosEstrada += metros;
      else metrosCidade += metros;
    });
    const totalMetrosClassificados = metrosCidade + metrosEstrada;
    const pctCidade = totalMetrosClassificados > 0 ? Math.round((metrosCidade / totalMetrosClassificados) * 100) : null;

    res.json({
      ok: true,
      distanciaTotalKm: Math.round(distanciaTotalKm * 10) / 10,
      pedagioTotalEstimado: Math.round(pedagioTotal * 100) / 100,
      duracaoIdaSegundos: parseInt(rota.duration) || null,
      pctCidadeEstimado: pctCidade,
    });
  } catch (err) {
    console.error('[calcular-distancia-deslocamento] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});
// ===== FIM CALCULADORA DE DESLOCAMENTO =====
// ===== PREÇO MÉDIO DE COMBUSTÍVEL — GOOGLE PLACES API =====
app.get('/preco-combustivel-medio', async (req, res) => {
  if (!GOOGLE_ROUTES_API_KEY) return res.status(500).json({ ok: false, erro: 'Chave da API do Google não configurada no servidor.' });

  try {
    const r = await fetch('https://places.googleapis.com/v1/places:searchText', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': GOOGLE_ROUTES_API_KEY,
        'X-Goog-FieldMask': 'places.displayName,places.fuelOptions',
      },
      body: JSON.stringify({
        textQuery: 'postos de gasolina em São Paulo, SP',
        maxResultCount: 20,
      }),
    });

    const d = await r.json();
    if (!r.ok) {
      console.error('[preco-combustivel-medio] erro Google:', JSON.stringify(d));
      return res.status(500).json({ ok: false, erro: d.error?.message || 'Erro ao consultar postos.' });
    }

    console.log('[preco-combustivel-medio] DEBUG total de lugares retornados=' + (d.places || []).length + ' | amostra=' + JSON.stringify((d.places || []).slice(0, 3)));

    const precosGasolina = [];
    const precosEtanol = [];

    (d.places || []).forEach(place => {
      const fuelPrices = place.fuelOptions?.fuelPrices || [];
      fuelPrices.forEach(fp => {
        const valor = Number(fp.price?.units || 0) + Number(fp.price?.nanos || 0) / 1e9;
        if (valor <= 0) return;
        if (fp.type === 'REGULAR_UNLEADED') precosGasolina.push(valor);
        if (fp.type === 'E100') precosEtanol.push(valor);
      });
    });

    const media = arr => arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : null;
    const mediaGasolina = media(precosGasolina);
    const mediaEtanol = media(precosEtanol);

    if (mediaGasolina === null && mediaEtanol === null) {
      return res.json({ ok: false, erro: 'Nenhum posto com preço de combustível disponível encontrado. Tente novamente mais tarde.' });
    }

    res.json({
      ok: true,
      precoGasolina: mediaGasolina !== null ? Math.round(mediaGasolina * 100) / 100 : null,
      precoEtanol: mediaEtanol !== null ? Math.round(mediaEtanol * 100) / 100 : null,
      amostraGasolina: precosGasolina.length,
      amostraEtanol: precosEtanol.length,
    });
  } catch (err) {
    console.error('[preco-combustivel-medio] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});
// ===== FIM PREÇO MÉDIO DE COMBUSTÍVEL =====

// ===== PREÇO MÉDIO DE COMBUSTÍVEL — PLANILHA OFICIAL DA ANP =====
function normalizarTexto(txt) {
  return (txt || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
}

app.get('/preco-combustivel-anp', async (req, res) => {
  try {
    const XLSX = require('xlsx');

    // 1. Acha o link da planilha mais recente na página de listagem da ANP
    const rListagem = await fetch('https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas');
    const htmlListagem = await rListagem.text();
    const match = htmlListagem.match(/https:\/\/www\.gov\.br\/anp\/pt-br\/assuntos\/precos-e-defesa-da-concorrencia\/precos\/arquivos-lpc\/\d{4}\/resumo_semanal_lpc[^"'\s)]+\.xlsx/i);
    if (!match) {
      console.error('[preco-combustivel-anp] não achou link da planilha na página da ANP.');
      return res.json({ ok: false, erro: 'Não foi possível localizar a planilha mais recente da ANP.' });
    }
    const urlPlanilha = match[0];
    console.log('[preco-combustivel-anp] DEBUG url planilha encontrada: ' + urlPlanilha);

    // 2. Baixa e faz o parse da planilha
    const rPlanilha = await fetch(urlPlanilha);
    if (!rPlanilha.ok) {
      console.error('[preco-combustivel-anp] erro ao baixar planilha: ' + rPlanilha.status);
      return res.json({ ok: false, erro: 'Erro ao baixar a planilha da ANP.' });
    }
    const bufferPlanilha = Buffer.from(await rPlanilha.arrayBuffer());
    const workbook = XLSX.read(bufferPlanilha, { type: 'buffer' });

    console.log('[preco-combustivel-anp] DEBUG abas encontradas: ' + JSON.stringify(workbook.SheetNames));

    // 3. A aba "ESTADOS" tem uma linha por Estado+Produto (formato longo), não colunas separadas por combustível.
    // Colunas: [0]data-ini [1]data-fim [2]ESTADO [3]MUNICIPIO(vazio aqui) [4]PRODUTO [5]nº postos [6]unidade [7]PRECO MEDIO REVENDA [8]desvio [9]min [10]max [11]coef.variação
    const nomeAbaEstados = workbook.SheetNames.find(n => normalizarTexto(n) === 'ESTADOS') || workbook.SheetNames[0];
    const sheet = workbook.Sheets[nomeAbaEstados];
    const linhas = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Busca sem depender de posição fixa de coluna: acha linhas onde ALGUMA célula é "SAO PAULO"
    const linhasSP = linhas.filter(linha => linha.some(cel => normalizarTexto(cel) === 'SAO PAULO'));
    console.log('[preco-combustivel-anp] DEBUG aba=' + nomeAbaEstados + ' | total linhas na aba=' + linhas.length + ' | linhas de SP encontradas=' + linhasSP.length + ' | amostra=' + JSON.stringify(linhasSP.slice(0, 6)));

    if (!linhasSP.length) {
      console.log('[preco-combustivel-anp] DEBUG primeiras 5 linhas da aba (pra conferir estrutura): ' + JSON.stringify(linhas.slice(0, 5)));
      return res.json({ ok: false, erro: 'Não foi possível localizar os dados de São Paulo na planilha da ANP.' });
    }

    let precoGasolina = null;
    let precoEtanol = null;
    linhasSP.forEach(linha => {
      // Acha qual célula da linha é o nome do produto (contém GASOLINA ou ETANOL)
      const idxProduto = linha.findIndex(cel => {
        const t = normalizarTexto(cel);
        return t.includes('GASOLINA') || t.includes('ETANOL');
      });
      if (idxProduto === -1) return;
      const produto = normalizarTexto(linha[idxProduto]);

      // O preço é o primeiro número plausível (entre 1 e 15) depois da célula do produto
      let preco = null;
      for (let i = idxProduto + 1; i < linha.length; i++) {
        const v = parseFloat(String(linha[i]).replace(',', '.'));
        if (!isNaN(v) && v > 1 && v < 15) { preco = v; break; }
      }
      if (preco === null) return;

      if (produto.includes('GASOLINA') && !produto.includes('ADITIVADA')) precoGasolina = preco;
      if (produto.includes('ETANOL')) precoEtanol = preco;
    });

    if (precoGasolina === null && precoEtanol === null) {
      return res.json({
        ok: false,
        erro: 'Encontrou linhas de São Paulo, mas não conseguiu identificar Gasolina/Etanol entre os produtos. Veja o log do servidor pra ajustar.',
      });
    }

    res.json({
      ok: true,
      precoGasolina,
      precoEtanol,
      fonte: 'ANP',
      urlPlanilha,
    });
  } catch (err) {
    console.error('[preco-combustivel-anp] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});
// ===== FIM PREÇO MÉDIO DE COMBUSTÍVEL (ANP) =====



// ===== PORTAL ADMIN — APROVAÇÕES DE ALUNAS =====

async function buscarAprovacaoPorId(pageId) {
  const r = await fetch('https://api.notion.com/v1/pages/' + pageId, {
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
  });
  const pagina = await r.json();
  if (!r.ok) return null;
  const p = pagina.properties;
  return {
    pageId: pagina.id,
    tipo: p['Tipo']?.select?.name || '',
    cpf: p['CPF Aluna']?.rich_text?.[0]?.plain_text || '',
    nome: p['Nome Aluna']?.rich_text?.[0]?.plain_text || '',
    contato: p['Contato']?.phone_number || '',
    modalidade: p['Modalidade']?.rich_text?.[0]?.plain_text || '',
    turmaAtual: p['Turma Atual']?.rich_text?.[0]?.plain_text || '',
    turmaNova: p['Turma Nova']?.rich_text?.[0]?.plain_text || '',
    planoAtual: p['Plano Atual']?.rich_text?.[0]?.plain_text || '',
    planoNovo: p['Plano Novo']?.rich_text?.[0]?.plain_text || '',
    frequenciaAtual: p['Frequência Atual']?.rich_text?.[0]?.plain_text || '',
    frequenciaNova: p['Frequência Nova']?.rich_text?.[0]?.plain_text || '',
    valorNovo: p['Valor Novo']?.number ?? null,
    motivo: p['Motivo']?.rich_text?.[0]?.plain_text || '',
    detalhesMulta: p['Detalhes Multa']?.rich_text?.[0]?.plain_text || '',
    status: p['Status']?.select?.name || '',
    matriculaPageId: (p['Matrícula']?.relation || [])[0]?.id || null,
  };
}

app.get('/portal-admin/aprovacoes', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + APROVACOES_ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sorts: [{ timestamp: 'created_time', direction: 'descending' }],
        page_size: 100,
      }),
    });
    const d = await r.json();
    if (!r.ok) {
      console.error('[portal-admin/aprovacoes] erro:', JSON.stringify(d));
      return res.status(500).json({ ok: false, erro: 'Erro ao consultar aprovações.' });
    }

    const pedidos = (d.results || []).map(pagina => {
      const p = pagina.properties;
      return {
        pageId: pagina.id,
        tipo: p['Tipo']?.select?.name || '',
        nome: p['Nome Aluna']?.rich_text?.[0]?.plain_text || '',
        cpf: p['CPF Aluna']?.rich_text?.[0]?.plain_text || '',
        contato: p['Contato']?.phone_number || '',
        modalidade: p['Modalidade']?.rich_text?.[0]?.plain_text || '',
        turmaAtual: p['Turma Atual']?.rich_text?.[0]?.plain_text || '',
        turmaNova: p['Turma Nova']?.rich_text?.[0]?.plain_text || '',
        planoAtual: p['Plano Atual']?.rich_text?.[0]?.plain_text || '',
        planoNovo: p['Plano Novo']?.rich_text?.[0]?.plain_text || '',
        frequenciaAtual: p['Frequência Atual']?.rich_text?.[0]?.plain_text || '',
        frequenciaNova: p['Frequência Nova']?.rich_text?.[0]?.plain_text || '',
        valorNovo: p['Valor Novo']?.number ?? null,
        motivo: p['Motivo']?.rich_text?.[0]?.plain_text || '',
        detalhesMulta: p['Detalhes Multa']?.rich_text?.[0]?.plain_text || '',
        status: p['Status']?.select?.name || '',
        criadoEm: pagina.created_time,
      };
    });

    res.json({ ok: true, pedidos });
  } catch (err) {
    console.error('[portal-admin/aprovacoes] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar aprovações.' });
  }
});

app.post('/portal-admin/aprovacoes/:id/aprovar', async (req, res) => {
  try {
    const pedido = await buscarAprovacaoPorId(req.params.id);
    if (!pedido) return res.status(404).json({ ok: false, erro: 'Pedido não encontrado.' });
    if (pedido.status !== 'Pendente') return res.status(400).json({ ok: false, erro: 'Este pedido já foi decidido.' });
    if (!pedido.matriculaPageId) return res.status(400).json({ ok: false, erro: 'Matrícula não vinculada a este pedido — verifique manualmente no Notion.' });

    let mensagemAluna = '';

    if (pedido.tipo === 'Mudança de Turma') {
      // Recheca vaga na hora de aprovar
      const rCheck = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { and: [
            { property: 'Modalidade', select: { equals: pedido.modalidade } },
            { property: 'Turma', select: { equals: pedido.turmaNova } },
            { property: 'Status', select: { equals: 'Ativa' } },
          ]},
          page_size: 200,
        }),
      });
      const dCheck = await rCheck.json();
      const dadosModalidade = MODALIDADES_MATRICULA[pedido.modalidade];
      const infoTurmaNova = dadosModalidade?.turmas.find(t => t.nome === pedido.turmaNova);
      const ocupadas = (dCheck.results || []).length;
      if (infoTurmaNova && ocupadas >= infoTurmaNova.limite) {
        return res.json({ ok: false, erro: 'A turma ' + pedido.turmaNova + ' encheu desde o pedido. Não é possível aprovar agora — negue ou oriente a aluna a escolher outro horário.' });
      }

      const propsAtualizacao = { 'Turma': { select: { name: pedido.turmaNova } } };
      if (infoTurmaNova?.dia) propsAtualizacao['Dia'] = { select: { name: infoTurmaNova.dia } };
      if (infoTurmaNova?.horario) propsAtualizacao['Horário'] = { select: { name: infoTurmaNova.horario } };
      if (infoTurmaNova?.professor) propsAtualizacao['Professor'] = { select: { name: infoTurmaNova.professor } };

      const rPatchTurma = await fetch('https://api.notion.com/v1/pages/' + pedido.matriculaPageId, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ properties: propsAtualizacao }),
      });
      if (!rPatchTurma.ok) {
        const eBody = await rPatchTurma.text();
        console.error('[aprovar mudanca-turma] Notion recusou atualizar matricula:', eBody);
        return res.status(500).json({ ok: false, erro: 'Notion recusou atualizar a matrícula — ver logs.' });
      }

      mensagemAluna = '✅ Sua mudança de turma foi aprovada! Nova turma: ' + pedido.turmaNova + '.';

    } else if (pedido.tipo === 'Mudança de Plano') {
      await fetch('https://api.notion.com/v1/pages/' + pedido.matriculaPageId, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          properties: {
            'Plano': { select: { name: pedido.planoNovo } },
            'Frequência': { select: { name: pedido.frequenciaNova } },
            'Valor': { number: pedido.valorNovo },
          },
        }),
      });

      mensagemAluna = '✅ Sua mudança de plano foi aprovada! Novo plano: ' + pedido.planoNovo + ' (' + pedido.frequenciaNova + ').';

    } else if (pedido.tipo === 'Cancelamento') {
      await fetch('https://api.notion.com/v1/pages/' + pedido.matriculaPageId, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ properties: { 'Status': { select: { name: 'Cancelado' } } } }),
      });

      mensagemAluna = '✅ Seu cancelamento foi confirmado. Qualquer dúvida sobre a multa ou aviso prévio, fale com a gente.';

    } else {
      return res.status(400).json({ ok: false, erro: 'Tipo de pedido desconhecido.' });
    }

    await fetch('https://api.notion.com/v1/pages/' + pedido.pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: { 'Status': { select: { name: 'Aprovado' } } } }),
    });

    if (pedido.contato) {
      try { await enviarWhatsApp(pedido.contato, mensagemAluna); } catch (e) { console.error('[aprovar] erro ao notificar aluna:', e.message); }
    }

    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-admin/aprovacoes/aprovar] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao aprovar.' });
  }
});

app.post('/portal-admin/aprovacoes/:id/negar', async (req, res) => {
  const motivoNegacao = req.body.motivoNegacao || '';
  try {
    const pedido = await buscarAprovacaoPorId(req.params.id);
    if (!pedido) return res.status(404).json({ ok: false, erro: 'Pedido não encontrado.' });
    if (pedido.status !== 'Pendente') return res.status(400).json({ ok: false, erro: 'Este pedido já foi decidido.' });

    await fetch('https://api.notion.com/v1/pages/' + pedido.pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Status': { select: { name: 'Negado' } },
          'Motivo Negação': { rich_text: [{ text: { content: motivoNegacao } }] },
        },
      }),
    });

    if (pedido.contato) {
      const msg = '❌ Seu pedido (' + pedido.tipo + ') não foi aprovado.' + (motivoNegacao ? ('\nMotivo: ' + motivoNegacao) : '') + '\n\nFale com a gente se quiser entender melhor.';
      try { await enviarWhatsApp(pedido.contato, msg); } catch (e) { console.error('[negar] erro ao notificar aluna:', e.message); }
    }

    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-admin/aprovacoes/negar] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao negar.' });
  }
});
// ===== FIM PORTAL ADMIN — APROVAÇÕES DE ALUNAS =====

// ===== PORTAL ADMIN — OCUPAÇÃO DE TURMAS =====
const CAPACIDADE_TURMAS_DB = 'cf832b7c91db44df8cb506d31b40bc3c';

app.get('/portal-admin/ocupacao', async (req, res) => {
  try {
    const rCap = await fetch('https://api.notion.com/v1/databases/' + CAPACIDADE_TURMAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ page_size: 100 }),
    });
    const dCap = await rCap.json();
    if (!rCap.ok) {
      console.error('[portal-admin/ocupacao] erro Capacidade:', JSON.stringify(dCap));
      return res.status(500).json({ ok: false, erro: 'Erro ao consultar Capacidade de Turmas.' });
    }

    const turmas = (dCap.results || []).map(pagina => {
      const p = pagina.properties;
      return {
        pageId: pagina.id,
        turma: p['Título']?.title?.[0]?.plain_text || '',
        modalidade: p['Modalidade']?.select?.name || '',
        turmaId: p['Turma ID']?.rich_text?.[0]?.plain_text || '',
        capacidadeRegular: p['Capacidade Regular']?.number ?? 0,
        vagaReposicaoExtra: p['Vaga Reposição Extra']?.number ?? 0,
      };
    });

    // Busca TODAS as alunas ativas de uma vez (mais eficiente que 1 query por turma)
    const rAlunas = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { or: [
          { property: 'Status', select: { equals: 'Ativa' } },
          { property: 'Status', select: { equals: 'Ativo' } },
        ]},
        page_size: 100,
      }),
    });
    const dAlunas = await rAlunas.json();
    const alunas = dAlunas.results || [];

    const DIA_ORDEM = { 'segunda': 1, 'terça': 2, 'terca': 2, 'quarta': 3, 'quinta': 4, 'sexta': 5, 'sábado': 6, 'sabado': 6, 'domingo': 7 };

    function extrairDiaEHora(turma) {
      const partes = (turma || '').trim().toLowerCase().split(/\s+/);
      const diaTexto = partes[0] || '';
      const dia = DIA_ORDEM[diaTexto] ?? 99;
      const horaMatch = (turma || '').match(/(\d{1,2})h(\d{2})?/i);
      const hora = horaMatch ? parseInt(horaMatch[1], 10) * 60 + (horaMatch[2] ? parseInt(horaMatch[2], 10) : 0) : 9999;
      return { dia, hora };
    }

    const resultado = turmas.map(t => {
      const ocupados = alunas.filter(pagina => {
        const p = pagina.properties;
        const modalidade = p['Modalidade']?.select?.name || '';
        const turma = p['Turma']?.select?.name || '';
        return modalidade === t.modalidade && turma === t.turma;
      }).length;

      return {
        ...t,
        capacidadeTotal: t.capacidadeRegular + t.vagaReposicaoExtra,
        ocupados,
      };
    });

    resultado.sort((a, b) => {
      const da = extrairDiaEHora(a.turma);
      const db = extrairDiaEHora(b.turma);
      if (da.dia !== db.dia) return da.dia - db.dia;
      return da.hora - db.hora;
    });

    res.json({ ok: true, turmas: resultado });
  } catch (err) {
    console.error('[portal-admin/ocupacao] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar ocupação.' });
  }
});

app.post('/portal-admin/ocupacao/salvar', async (req, res) => {
  const { pageId, capacidadeRegular, vagaReposicaoExtra } = req.body;

  if (!pageId) {
    return res.status(400).json({ ok: false, erro: 'pageId é obrigatório.' });
  }

  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Capacidade Regular': { number: Number(capacidadeRegular) },
          'Vaga Reposição Extra': { number: Number(vagaReposicaoExtra) },
        },
      }),
    });
    const d = await r.json();
    if (!r.ok) {
      console.error('[portal-admin/ocupacao/salvar] erro:', JSON.stringify(d));
      return res.status(500).json({ ok: false, erro: 'Erro ao salvar.' });
    }
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-admin/ocupacao/salvar] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao salvar.' });
  }
});
// ===== FIM PORTAL ADMIN — OCUPAÇÃO DE TURMAS =====

// ===== PORTAL ADMIN — CADASTRO DE PROFESSORES (CRUD) =====
app.get('/portal-admin/professores/:pageId', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + req.params.pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const pagina = await r.json();
    if (!r.ok) return res.status(404).json({ ok: false, erro: 'Professor não encontrado.' });

    const p = pagina.properties;
    res.json({
      ok: true,
      professor: {
        pageId: pagina.id,
        nome: p['Nome']?.title?.[0]?.plain_text || '',
        cpf: p['CPF']?.rich_text?.[0]?.plain_text || '',
        telefone: p['Telefone']?.phone_number || '',
        email: p['Email']?.email || '',
        endereco: p['Endereço']?.rich_text?.[0]?.plain_text || '',
        modalidades: (p['Modalidades']?.multi_select || []).map(m => m.name),
        valorHora: p['Valor Hora']?.number ?? null,
      },
    });
  } catch (err) {
    console.error('[portal-admin/professores/:pageId] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao buscar professor.' });
  }
});

app.post('/portal-admin/professores/salvar', async (req, res) => {
  const { pageId, nome, cpf, telefone, email, endereco, modalidades, valorHora } = req.body;

  if (!nome || !String(nome).trim()) {
    return res.status(400).json({ ok: false, erro: 'Nome é obrigatório.' });
  }

  const properties = {
    Nome: { title: [{ text: { content: nome } }] },
    CPF: { rich_text: [{ text: { content: cpf || '' } }] },
    Telefone: { phone_number: telefone || null },
    Email: { email: email || null },
    Endereço: { rich_text: [{ text: { content: endereco || '' } }] },
    Modalidades: { multi_select: (modalidades || []).map(m => ({ name: m })) },
    'Valor Hora': { number: valorHora !== undefined && valorHora !== null && valorHora !== '' ? Number(valorHora) : null },
  };

  try {
    if (pageId) {
      // Atualiza professor existente
      const r = await fetch('https://api.notion.com/v1/pages/' + pageId, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ properties }),
      });
      const d = await r.json();
      if (!r.ok) {
        console.error('[portal-admin/professores/salvar] erro update:', JSON.stringify(d));
        return res.status(500).json({ ok: false, erro: 'Erro ao atualizar professor.' });
      }
      return res.json({ ok: true, pageId: d.id, criado: false });
    } else {
      // Cria professor novo
      const r = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ parent: { database_id: PROFESSORES_CADASTRO_DB }, properties }),
      });
      const d = await r.json();
      if (!r.ok) {
        console.error('[portal-admin/professores/salvar] erro create:', JSON.stringify(d));
        return res.status(500).json({ ok: false, erro: 'Erro ao criar professor.' });
      }
      return res.json({ ok: true, pageId: d.id, criado: true });
    }
  } catch (err) {
    console.error('[portal-admin/professores/salvar] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao salvar professor.' });
  }
});
// ===== FIM PORTAL ADMIN — CADASTRO DE PROFESSORES =====

app.post('/portal/login/solicitar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  try {
    const prof = await buscarProfessorPorCpf(cpfLimpo);
    if (!prof) return res.json({ ok: false, erro: 'CPF não encontrado. Fale com o Fábio.' });
    await enviarOtp('prof_' + cpfLimpo, prof.telefone, prof.nome);
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal/login/solicitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal/login/verificar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  const verificacao = verificarOtp('prof_' + cpfLimpo, req.body.codigo);
  if (!verificacao.ok) return res.json(verificacao);
  try {
    const prof = await buscarProfessorPorCpf(cpfLimpo);
    if (!prof) return res.json({ ok: false, erro: 'Cadastro não encontrado.' });
    const dadosLogin = { nome: prof.nome, telefone: prof.telefone, endereco: prof.endereco, email: prof.email, cpf: cpfLimpo };
    const sessionToken = criarSessao('prof_' + cpfLimpo, dadosLogin);
    res.json({ ok: true, ...dadosLogin, sessionToken });
  } catch (err) {
    console.error('[portal/login/verificar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal/sessao/verificar', (req, res) => {
  const dados = verificarSessao(req.body.sessionToken);
  if (!dados) return res.json({ ok: false, erro: 'Sessão expirada.' });
  res.json({ ok: true, ...dados });
});

app.post('/portal/atualizar-cadastro', async (req, res) => {
  const { cpf, nome, telefone, endereco, email, sessionToken } = req.body;
  const cpfLimpo = (cpf || '').replace(/\D/g, '');
  if (!cpfLimpo || !nome || !telefone) {
    return res.status(400).json({ ok: false, erro: 'Preencha nome e telefone.' });
  }
  const dadosSessao = verificarSessao(sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessao.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const prof = await buscarProfessorPorCpf(cpfLimpo);
    if (!prof) return res.status(404).json({ ok: false, erro: 'Cadastro não encontrado.' });

    await fetch('https://api.notion.com/v1/pages/' + prof.pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Nome': { title: [{ text: { content: nome } }] },
          'Telefone': { phone_number: telefone },
          'Endereço': { rich_text: [{ text: { content: endereco || '' } }] },
          'Email': { email: email || null },
        },
      }),
    });
    try { await enviarWhatsApp(WHATSAPP_FABIO, 'ℹ️ *Professor(a) atualizou o cadastro* — ' + nome + ' (' + telefone + ')'); } catch(e) {}
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal/atualizar-cadastro] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

async function turmasDoProfessor(nome) {
  const encontradas = [];
  const cacheProfessorPorModalidade = {};
  for (const [modalidade, dados] of Object.entries(MODALIDADES_MATRICULA)) {
    if (!cacheProfessorPorModalidade[modalidade]) {
      try {
        cacheProfessorPorModalidade[modalidade] = await buscarProfessorDaTurma(modalidade);
      } catch (e) {
        console.error('[turmasDoProfessor] erro ao buscar professor no Notion, usando fallback fixo:', e.message);
        cacheProfessorPorModalidade[modalidade] = {};
      }
    }
    const mapaProfessores = cacheProfessorPorModalidade[modalidade];
    dados.turmas.forEach(t => {
      // Prioriza o(s) professor(es) cadastrado(s) no Notion (multi_select); se a
      // turma ainda nao estiver migrada la (lista vazia), usa o valor fixo do
      // codigo como fallback seguro.
      const professoresAtual = mapaProfessores[t.nome] || [t.professor];
      if (professoresAtual.includes(nome)) encontradas.push({ modalidade, ...t, professor: professoresAtual[0], professores: professoresAtual });
    });
  }
  return encontradas;
}

app.get('/portal/turmas/:nome', async (req, res) => {
  const nome = decodeURIComponent(req.params.nome);
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessao.nome !== nome) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const turmasEncontradas = await turmasDoProfessor(nome);
    const resultado = [];
    for (const t of turmasEncontradas) {
      const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filter: { and: [
            { property: 'Modalidade', select: { equals: t.modalidade } },
            { property: 'Turma', select: { equals: t.nome } },
            { property: 'Status', select: { equals: 'Ativa' } },
          ]},
          page_size: 200,
        }),
      });
      const d = await r.json();
      const alunos = (d.results || []).map(pagina => {
        const pp = pagina.properties;
        return {
          nome: pp['Nome']?.title?.[0]?.plain_text || '',
          telefone: (pp['Contato']?.phone_number || '').replace(/\D/g, ''),
        };
      }).filter(a => a.nome);
      const capacidadeReal = await buscarCapacidadeDaTurma(t.modalidade);
      const limiteReal = capacidadeReal[t.nome] ?? t.limite;
      resultado.push({ modalidade: t.modalidade, turma: t.nome, dia: t.dia, horario: t.horario, alunasAtivas: alunos.length, limite: limiteReal, alunos });
    }
    res.json({ ok: true, turmas: resultado });
  } catch (err) {
    console.error('[portal] erro ao buscar turmas:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.get('/portal/feriados/:nome', async (req, res) => {
  const nome = decodeURIComponent(req.params.nome);
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessao.nome !== nome) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const diasDoProfessor = new Set((await turmasDoProfessor(nome)).map(t => t.dia));
    if (diasDoProfessor.size === 0) return res.json({ ok: true, feriados: [] });

    const hoje = new Date();
    const anoAtual = hoje.getFullYear();
    const feriadosAno = await getFeriadosDoAno(anoAtual);
    const feriadosProxAno = hoje.getMonth() >= 9 ? await getFeriadosDoAno(anoAtual + 1) : new Set();
    const todosFeriados = new Set([...feriadosAno, ...feriadosProxAno]);

    const nomesDias = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
    const em90dias = new Date(hoje.getTime() + 90 * 24 * 60 * 60000);
    const resultado = [];

    todosFeriados.forEach(dataStr => {
      const data = new Date(dataStr + 'T12:00:00-03:00');
      if (data < hoje || data > em90dias) return;
      const diaSemana = nomesDias[data.getDay()];
      if (diasDoProfessor.has(diaSemana)) resultado.push({ data: dataStr, diaSemana });
    });

    resultado.sort((a, b) => a.data.localeCompare(b.data));
    res.json({ ok: true, feriados: resultado });
  } catch (err) {
    console.error('[portal] erro ao buscar feriados:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

const VALOR_AULA_PROFESSOR = {
  'Gustra': 80,
  'Titzi': 80,
  'Guilherme': 80,
};

app.get('/portal/rendimento/:nome', async (req, res) => {
  const nome = decodeURIComponent(req.params.nome);
  const dadosSessao = verificarSessao(req.query.sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessao.nome !== nome) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  const valorAula = VALOR_AULA_PROFESSOR[nome];
  if (!valorAula) return res.json({ ok: true, disponivel: false });

  try {
    const turmas = await turmasDoProfessor(nome);
    if (turmas.length === 0) return res.json({ ok: true, disponivel: false });

    const hoje = new Date();
    const ano = hoje.getFullYear();
    const mes = hoje.getMonth();
    const ultimoDia = new Date(ano, mes + 1, 0).getDate();
    const nomesDias = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];

    const rDecisoes = await fetch('https://api.notion.com/v1/databases/' + DECISOES_FERIADO_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Professor', rich_text: { equals: nome } }, page_size: 100 }),
    });
    const dDecisoes = await rDecisoes.json();
    const decisaoPorData = {};
    (dDecisoes.results || []).forEach(p => {
      const data = p.properties['Data do Feriado']?.date?.start;
      const decisao = p.properties['Decisão']?.select?.name;
      if (data) decisaoPorData[data] = decisao;
    });

    const feriadosDoAno = await getFeriadosDoAno(ano);

    let totalAulas = 0;
    const detalhes = [];

    for (const turma of turmas) {
      let contagem = 0;
      for (let dia = 1; dia <= ultimoDia; dia++) {
        const dataObj = new Date(ano, mes, dia);
        const diaSemanaNome = nomesDias[dataObj.getDay()];
        if (diaSemanaNome !== turma.dia) continue;
        const dataStr = ano + '-' + String(mes + 1).padStart(2, '0') + '-' + String(dia).padStart(2, '0');
        if (feriadosDoAno.has(dataStr) && decisaoPorData[dataStr] !== 'Mantém a aula') continue;
        contagem++;
      }
      totalAulas += contagem;
      detalhes.push({ modalidade: turma.modalidade, turma: turma.nome, aulas: contagem });
    }

    res.json({
      ok: true, disponivel: true, valorAula, totalAulas,
      totalPrevisto: totalAulas * valorAula, detalhes,
      mesAno: (mes + 1) + '/' + ano,
    });
  } catch (err) {
    console.error('[portal] erro ao calcular rendimento:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal/feriado-resposta', async (req, res) => {
  const { nome, data, decisao, sessionToken } = req.body;
  if (!nome || !data || !decisao) return res.status(400).json({ ok: false, erro: 'Campos obrigatórios faltando.' });
  const dadosSessao = verificarSessao(sessionToken);
  if (!dadosSessao) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessao.nome !== nome) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: DECISOES_FERIADO_DB },
        properties: {
          'Título': { title: [{ text: { content: nome + ' — ' + data } }] },
          'Professor': { rich_text: [{ text: { content: nome } }] },
          'Data do Feriado': { date: { start: data } },
          'Decisão': { select: { name: decisao } },
        },
      }),
    });
    const dataFmt = data.split('-').reverse().join('/');
    const msgFabio = '🗓️ *Decisão de feriado* — ' + nome + '\nData: ' + dataFmt + '\nDecisão: ' + decisao;
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgFabio); } catch(e) {}
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal] erro ao salvar decisao:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// PROPOSTA APROVADA -> cria Apresentacao automaticamente
// ============================================================
const PROPOSTAS_DB = '2c6c45031f73804f8f90e6e7439d7e1c';
const APRESENTACOES_DB_PARA_WEBHOOK_PROPOSTA = '2b9c45031f7380828d34f47353b066e7';

app.post('/webhook-proposta-aprovada', async (req, res) => {
  res.status(200).json({ ok: true });
  try {
    const body = req.body || {};
    const pageId = (body.data && body.data.id) || body.pageId || body.page_id || null;
    if (!pageId) { console.error('[webhook-proposta-aprovada] sem page id.'); return; }

    const rProposta = await fetch('https://api.notion.com/v1/pages/' + pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const proposta = await rProposta.json();
    const p = proposta.properties;

    const status = p['Status']?.status?.name;
    if (status !== 'Aprovado - Aguardando contrato') {
      console.log('[webhook-proposta-aprovada] status nao e Aprovado, ignorando: ' + status);
      return;
    }

    // Ja existe Apresentacao vinculada? Evita duplicar se o automation disparar de novo.
    // Consulta direto na planilha de Apresentacoes (a Proposta nao guarda mais esse link).
    const rCheckExistente = await fetch('https://api.notion.com/v1/databases/' + APRESENTACOES_DB_PARA_WEBHOOK_PROPOSTA + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Proposta', relation: { contains: pageId } }, page_size: 1 }),
    });
    const dCheckExistente = await rCheckExistente.json();
    if ((dCheckExistente.results || []).length > 0) {
      const apresentacaoExistenteId = dCheckExistente.results[0].id;
      console.log('[webhook-proposta-aprovada] ja tem apresentacao vinculada, sincronizando calendario mesmo assim: ' + apresentacaoExistenteId);
      try { await sincronizarApresentacaoComCalendar(apresentacaoExistenteId); } catch (e) { console.error('[webhook-proposta-aprovada] erro ao sincronizar calendario (apresentacao existente):', e.message); }
      return;
    }

    const enderecoProp = p['Endereço'] || null; // tipo "place" - repassado como veio, sem remontar a estrutura
    const tituloDaProposta = p['Local']?.title?.[0]?.plain_text || '';
    const contratanteNomes = (p['Contratante']?.multi_select || []).map(o => o.name);
    const tituloApresentacao = tituloDaProposta || contratanteNomes.join(', ') || 'Apresentação sem local definido';

    const dataApresentacao = p['Data']?.date?.start || null;
    const horarioApresentacaoProposta = p['Horário Apresentação']?.rich_text?.[0]?.plain_text || '';
    const trabalhosIds = (p['🎭 Trabalhos']?.relation || []).map(r => r.id);
    const integrantesIds = (p['Integrantes']?.relation || []).map(r => r.id);

    const propsNovaApresentacao = {
      'LOCAL': { title: [{ text: { content: tituloApresentacao } }] },
      'Proposta': { relation: [{ id: pageId }] },
    };
    if (enderecoProp?.place) propsNovaApresentacao['Endereço'] = { place: enderecoProp.place };
    if (dataApresentacao) propsNovaApresentacao['Data da Apresentação'] = { date: { start: dataApresentacao } };
    if (horarioApresentacaoProposta) propsNovaApresentacao['Horário Apresentação'] = { rich_text: [{ text: { content: horarioApresentacaoProposta } }] };
    if (trabalhosIds.length) propsNovaApresentacao['🎭 Trabalhos'] = { relation: trabalhosIds.map(id => ({ id })) };
    if (integrantesIds.length) propsNovaApresentacao['ELENCO'] = { relation: integrantesIds.map(id => ({ id })) };

    const rCriar = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: APRESENTACOES_DB_PARA_WEBHOOK_PROPOSTA },
        properties: propsNovaApresentacao,
      }),
    });
    if (!rCriar.ok) { const t = await rCriar.text(); throw new Error('Notion criar apresentacao: ' + t); }
    const novaApresentacao = await rCriar.json();

    const msgFabio = '🎉 *Proposta aprovada — Apresentação criada automaticamente*\n\nLocal: ' + tituloApresentacao + (dataApresentacao ? ('\nData: ' + dataApresentacao.split('-').reverse().join('/')) : '\nData: (não definida)');
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgFabio); } catch(e) {}

    console.log('[webhook-proposta-aprovada] apresentacao criada: ' + novaApresentacao.id);

    try {
      await sincronizarApresentacaoComCalendar(novaApresentacao.id);
    } catch (e) {
      console.error('[webhook-proposta-aprovada] erro ao sincronizar calendario (apresentacao nova):', e.message);
    }
  } catch (err) {
    console.error('[webhook-proposta-aprovada] erro:', err.message);
  }
});

// ============================================================
// CONTRATOS DE DOCENCIA — PF-RPA / PJ-MEI / PJ-ME (v2)
// Modelo de remuneracao ramificado pela titularidade: hora/aula OU percentual.
// Dados completos do prestador persistidos como JSON, para regeneracao fiel do texto.
// ============================================================
const CONTRATOS_PROF_DB = 'f1b934a8100143019ac7f5f877c405f1';

function calcularIdade(dataNascISO) {
  if (!dataNascISO) return null;
  const hoje = new Date();
  const nasc = new Date(dataNascISO + 'T12:00:00-03:00');
  let idade = hoje.getFullYear() - nasc.getFullYear();
  const m = hoje.getMonth() - nasc.getMonth();
  if (m < 0 || (m === 0 && hoje.getDate() < nasc.getDate())) idade--;
  return idade;
}

// Biblioteca de clausulas comuns, texto integral (nao resumido)
function clausulasComuns345(parteContratada) {
  return `CLÁUSULA 3ª — FORMA DE EXECUÇÃO E AUTONOMIA
3.1. A/O ${parteContratada} organiza livremente a condução técnica e pedagógica de suas aulas, não estando sujeito(a) a registro de ponto, ordens de serviço ou fiscalização de método de trabalho por parte da CONTRATANTE.
3.2. A CONTRATANTE limita-se a disponibilizar espaço e equipamentos adequados e seguros, cabendo-lhe apenas a organização operacional (reserva de sala, comunicação com alunos, agenda geral do espaço), sem qualquer ingerência sobre o conteúdo pedagógico.
3.3. A/O ${parteContratada} pode recusar datas propostas e sugerir remanejamentos de agenda, sem que isso configure penalidade, advertência ou qualquer forma de sanção disciplinar.

CLÁUSULA 4ª — LOCAL DE PRESTAÇÃO DOS SERVIÇOS
4.1. Os serviços serão prestados nas dependências do Espaço Liquidificador, situado na Rua Dr. Carvalho de Mendonça, 67, Campos Elíseos, São Paulo/SP, podendo as partes acordar, excepcionalmente, outro local de comum acordo.

CLÁUSULA 5ª — AGENDA DE AULAS
5.1. As aulas serão ministradas conforme agenda pactuada de comum acordo entre as partes (Anexo I deste instrumento).
5.2. A agenda constitui estimativa de escopo do serviço contratado, não configurando jornada de trabalho nem obrigação de disponibilidade contínua; alterações de agenda podem ser feitas por simples acordo entre as partes.
5.3. A impossibilidade eventual de ministrar determinada aula deverá ser comunicada com a antecedência possível nas circunstâncias, sem que isso caracterize falta disciplinar ou indício de subordinação.`;
}

function clausula9(parteContratada) {
  return `CLÁUSULA 9ª — OBRIGAÇÕES DA CONTRATANTE
A CONTRATANTE obriga-se a: a) disponibilizar espaço e equipamentos em condições adequadas e seguras de uso; b) efetuar o pagamento devido contra a respectiva nota fiscal ou recibo, nos prazos pactuados; c) fornecer as informações operacionais necessárias à boa condução das aulas (calendário de feriados, ocorrências relevantes no espaço); d) manter a manutenção preventiva e corretiva das instalações e equipamentos.`;
}

function clausula10(parteContratada) {
  return `CLÁUSULA 10ª — SEGURANÇA E RESPONSABILIDADE CIVIL
10.1. A CONTRATANTE responde pela segurança estrutural do espaço e pela integridade e manutenção dos equipamentos disponibilizados.
10.2. A/O ${parteContratada} responde pela adequação técnica e pedagógica da condução das aulas, e por eventuais danos decorrentes de conduta própria ou de eventual substituto por ela/ele indicado.
10.3. A responsabilidade por qualquer incidente será apurada conforme sua causa real: defeito estrutural ou de equipamento imputa-se à CONTRATANTE; falha de conduta técnica imputa-se à/ao ${parteContratada}. Fica vedada qualquer renúncia genérica de direitos por qualquer das partes.`;
}

function clausula11(parteContratada) {
  return `CLÁUSULA 11ª — NÃO-EXCLUSIVIDADE E INDEPENDÊNCIA ECONÔMICA
11.1. O presente contrato não implica exclusividade, sendo livre à/ao ${parteContratada} prestar serviços de mesma natureza a quaisquer terceiros, concorrentes ou não da CONTRATANTE.
11.2. A/O ${parteContratada} declara exercer atividade autônoma e independente, possuir outras fontes de receita, e que a presente prestação de serviços tem caráter acessório em sua atividade econômica.
11.3. As partes reconhecem expressamente a inexistência de dependência econômica entre si.`;
}

function clausula12(parteContratada) {
  return `CLÁUSULA 12ª — PROPRIEDADE INTELECTUAL
12.1. O método, a metodologia e o material didático desenvolvidos pela/pelo ${parteContratada} permanecem de sua exclusiva titularidade, sendo autorizado à CONTRATANTE seu uso restrito à realização das aulas objeto deste contrato, e apenas durante a vigência contratual.`;
}

function clausula13(parteContratada, publicoMenor) {
  return `CLÁUSULA 13ª — PROTEÇÃO DE DADOS PESSOAIS (LGPD)
13.1. As partes observarão a Lei nº 13.709/2018 (LGPD) em todo o tratamento de dados pessoais relacionado à execução deste contrato.
13.2. Quanto aos dados dos alunos, a CONTRATANTE figura como controladora, e a/o ${parteContratada} trata tais dados apenas na medida necessária à ministração das aulas, observando sigilo e confidencialidade.
13.3. Os dados pessoais da/do titular da parte CONTRATADA são tratados pela CONTRATANTE para fins de execução contratual e cumprimento de obrigações legais.${publicoMenor ? '\n13.4. Havendo alunos menores de idade nas turmas, o tratamento de seus dados observará o disposto no art. 14 da LGPD, conforme detalhado na Cláusula 14-A deste contrato.' : ''}`;
}

function clausula14(parteContratada) {
  return `CLÁUSULA 14ª — CONFIDENCIALIDADE
14.1. As partes comprometem-se a manter sigilo sobre quaisquer informações confidenciais a que tenham acesso em razão deste contrato, obrigação que subsiste mesmo após o término da relação contratual.`;
}

function clausula16() {
  return `CLÁUSULA 16ª — DISPOSIÇÕES GERAIS
16.1. Qualquer alteração deste contrato somente terá validade se formalizada por escrito, mediante termo aditivo assinado por ambas as partes.
16.2. A tolerância de qualquer das partes quanto ao descumprimento de qualquer cláusula não implicará novação, renúncia ou alteração do pactuado, podendo a parte exigir o cumprimento a qualquer tempo.
16.3. As partes atuam como entes independentes, respondendo cada qual, com exclusividade, por suas próprias obrigações legais, fiscais, previdenciárias e trabalhistas, inexistindo solidariedade entre elas.
16.4. A eventual nulidade de qualquer cláusula não prejudicará as demais, que permanecerão em vigor, comprometendo-se as partes a substituir a disposição afetada por outra de efeito equivalente.`;
}

function clausula17() {
  return `CLÁUSULA 17ª — FORO
17.1. As partes elegem o foro da Comarca de São Paulo/SP para dirimir quaisquer controvérsias oriundas deste contrato.`;
}

function anexoIAgenda(agenda) {
  return `ANEXO I — AGENDA DE AULAS
Integra este contrato a seguinte agenda, pactuada de comum acordo entre as partes, conforme referida na Cláusula 5ª:

${agenda && agenda.trim() ? agenda.trim() : 'A ser detalhada e formalizada por escrito entre as partes, podendo ser ajustada a qualquer tempo mediante simples acordo, sem necessidade de aditivo contratual.'}`;
}

function montarTextoContratoProfessor(d) {
  const {
    trilha, modalidadeLabel, titularidade, formato, valorHora, periodicidade, prazoQuitacao,
    substituicao, compensacaoCancelamento, compensacaoHoras, compensacaoPct,
    vigencia, vigenciaInicio, vigenciaFim, avisoPrevioDias, publicoMenor,
    dataNascimento, repLegNome, repLegCpf, repLegParentesco,
    nome, nomeSocial, nacionalidade, estadoCivil, rg, cpf, nitPis, endereco,
    razaoSocial, cnpj, ocupacao, titularNome, titularCpf,
    cnae, repLegal, repCpf,
    contatoEmergenciaNome, contatoEmergenciaTelefone,
    pctProfessor, pctEspaco, baseCalculo, quemFatura,
    agenda,
  } = d;

  // Titularidade == professor traz alunos: nao gera contrato de docencia
  if (titularidade === 'Professor traz alunos') {
    return { bloqueado: true, motivo: 'Neste modelo, o(a) professor(a) utiliza o espaço mediante aluguel (cessão onerosa de uso), e não um contrato de docência. Utilize o instrumento de Cessão Onerosa de Uso para este caso.' };
  }

  const idadePrestador = calcularIdade(dataNascimento);
  const ehPrestadorMenor = idadePrestador !== null && idadePrestador < 18;
  const ehOficina = formato === 'Oficina com prazo';
  const ehPF = trilha === 'PF-RPA';
  const ehMEI = trilha === 'PJ-MEI';
  const ehPercentual = titularidade === 'Misto';
  const parteContratada = ehPF ? 'CONTRATADO(A)' : 'CONTRATADA';
  const docFiscal = ehPF ? 'RPA' : 'NFS-e';
  const nomeExibicaoPrestador = ehPF ? (nomeSocial || nome) : (nomeSocial || razaoSocial);

  let titulo, qualificacao;
  if (ehMEI) {
    titulo = 'CONTRATO DE PRESTAÇÃO DE SERVIÇOS DE DOCÊNCIA\nRegime: prestação de serviços entre pessoas jurídicas — CONTRATADA constituída como Microempreendedor Individual (MEI)';
    qualificacao = `CONTRATADA: ${razaoSocial}${nomeSocial ? ' ("' + nomeSocial + '")' : ''}, microempreendedor individual inscrito no CNPJ nº ${cnpj}, sede na ${endereco}, atuando na ocupação de ${ocupacao}, representada por seu(sua) titular ${titularNome}, CPF nº ${titularCpf}, doravante CONTRATADA.`;
  } else if (trilha === 'PJ-ME') {
    titulo = 'CONTRATO DE PRESTAÇÃO DE SERVIÇOS DE DOCÊNCIA\nRegime: prestação de serviços entre pessoas jurídicas — CONTRATADA constituída como Microempresa (ME)';
    qualificacao = `CONTRATADA: ${razaoSocial}${nomeSocial ? ' ("' + nomeSocial + '")' : ''}, microempresa inscrita no CNPJ nº ${cnpj}, sede na ${endereco}, CNAE ${cnae}, representada por seu(sua) sócio(a)/administrador(a) ${repLegal}, CPF nº ${repCpf}, doravante CONTRATADA.`;
  } else {
    titulo = 'CONTRATO DE PRESTAÇÃO DE SERVIÇOS AUTÔNOMOS DE DOCÊNCIA\nRegime: prestação de serviços autônomos por pessoa física (RPA)';
    qualificacao = `CONTRATADO(A): ${nome}${nomeSocial ? ' ("' + nomeSocial + '")' : ''}, ${nacionalidade || ''}, ${estadoCivil || ''}, RG nº ${rg}, CPF nº ${cpf}, inscrito(a) no INSS como contribuinte individual sob NIT/PIS nº ${nitPis || ''}, residente na ${endereco}, doravante CONTRATADO(A).`;
  }

  const clausula1 = `CLÁUSULA 1ª — OBJETO
1.1. Prestação${ehPF ? ' de serviços autônomos de docência' : (', pela CONTRATADA, de serviços de docência')} na modalidade ${modalidadeLabel}, nas dependências do Espaço Liquidificador.${ehOficina ? ' A presente oferta está estruturada como OFICINA, de prazo determinado.' : ''}
1.2. Os serviços serão prestados com plena autonomia técnica, pedagógica e metodológica, cabendo exclusivamente à/ao ${parteContratada} definir o conteúdo, o plano de aula e a condução das atividades, sem qualquer ingerência da CONTRATANTE quanto ao modo de execução.`;

  const clausula2 = ehPF
    ? `CLÁUSULA 2ª — NATUREZA AUTÔNOMA DA RELAÇÃO
2.1. O presente instrumento estabelece relação civil de prestação de serviços autônomos, sem qualquer vínculo empregatício entre as partes, nos termos do art. 442-B da CLT e dos arts. 593 e seguintes do Código Civil.
2.2. Os serviços são executados por conta própria da/do CONTRATADO(A), sem subordinação jurídica, sem controle de jornada, sem sujeição a poder disciplinar da CONTRATANTE, e sem exclusividade.
2.3. A/O CONTRATADO(A) assume integralmente seus próprios tributos, contribuições previdenciárias e obrigações fiscais decorrentes desta prestação de serviços.`
    : `CLÁUSULA 2ª — NATUREZA DA RELAÇÃO
2.1. O presente instrumento estabelece relação civil de prestação de serviços entre pessoas jurídicas independentes, sem qualquer vínculo empregatício, societário ou de subordinação entre as partes, nos termos do art. 442-B da CLT e dos arts. 593 e seguintes do Código Civil.
2.2. A CONTRATADA executa os serviços por sua própria conta e risco, sem subordinação jurídica à CONTRATANTE, sem controle de jornada, sem sujeição a poder disciplinar, e sem exclusividade.
2.3. Cada parte responde exclusivamente por suas próprias obrigações legais, fiscais, previdenciárias e trabalhistas, inexistindo solidariedade entre elas.`;

  const clausulas3a5 = clausulasComuns345(parteContratada);

  let clausula6 = '';
  if (substituicao !== 'Não Permitida') {
    const varianteAviso = substituicao === 'Livre'
      ? 'independe de autorização prévia da CONTRATANTE, bastando comunicação com antecedência razoável'
      : 'depende de aviso prévio à CONTRATANTE, com antecedência razoável';
    clausula6 = `CLÁUSULA 6ª — FACULDADE DE SUBSTITUIÇÃO
6.1. É facultado à/ao ${parteContratada}, sempre que lhe convier, fazer-se substituir na ministração das aulas por (a) profissional que já preste serviços ao Espaço na mesma modalidade, ou (b) profissional externo devidamente habilitado.
6.2. A substituição ${varianteAviso}.
6.3. A/O ${parteContratada} garante a habilitação técnica e a aptidão do(a) substituto(a) indicado(a), respondendo por sua idoneidade. Em modalidades de risco, o(a) substituto(a) deverá atender ao mesmo padrão de aptidão documentado exigido neste contrato.
6.4. A remuneração correspondente é devida a quem efetivamente ministrar a aula, mediante emissão do respectivo ${docFiscal} por essa pessoa.
6.5. A ampla faculdade de substituição ora prevista evidencia a inexistência de pessoalidade na prestação dos serviços.`;
  }

  const compensacaoTexto = compensacaoCancelamento
    ? `\n7.6. Em caso de cancelamento por iniciativa exclusiva da CONTRATANTE com antecedência inferior a ${compensacaoHoras} horas do horário previsto, será devido à/ao ${parteContratada} o equivalente a ${compensacaoPct}% do valor da hora/aula cancelada.`
    : '';

  let clausula7;
  if (ehPercentual) {
    const baseCalculoLabel = baseCalculo === 'bruta' ? 'receita bruta' : 'receita líquida';
    const alertaBaseBruta = baseCalculo === 'bruta'
      ? '\n7.2-A. Ficam as partes cientes de que, ao se adotar a receita bruta como base de cálculo, a margem líquida efetiva da CONTRATANTE tende a ficar abaixo do percentual nominal a ela atribuído, uma vez que os tributos incidentes sobre a fatia repassada à/ao CONTRATADO(A) permanecem a cargo da CONTRATANTE.'
      : '';
    const faturamentoTexto = quemFatura === 'professor'
      ? `O faturamento perante o aluno pagante é realizado diretamente pela/pelo própria/próprio ${parteContratada}, que repassa à CONTRATANTE o percentual de ${pctEspaco}% correspondente.`
      : `A CONTRATANTE responsabiliza-se pela matrícula, cobrança e faturamento dos alunos pagantes, apurando e repassando à/ao ${parteContratada} o percentual devido.`;
    const retencaoPorTrilha = ehPF
      ? `Sobre o valor do repasse recebido pela/pelo CONTRATADO(A) incidem as retenções legais aplicáveis ao contribuinte individual (INSS 11%, até o teto, e IRRF conforme tabela progressiva quando aplicável), cabendo à CONTRATANTE, ainda, o recolhimento da contribuição patronal (20%) sobre tal remuneração.`
      : ehMEI
      ? `O repasse é pago contra emissão de NFS-e pela CONTRATADA, sem retenção previdenciária patronal, por se tratar de MEI em atividade de docência, ressalvada obrigação legal superveniente.`
      : `Sobre o repasse incidem as retenções tributárias aplicáveis (ISS, IR, CSLL, PIS, COFINS, INSS quando aplicável) conforme o regime tributário da CONTRATADA e a legislação municipal vigente, cabendo à CONTRATADA validar tais retenções junto à sua contabilidade.`;
    clausula7 = `CLÁUSULA 7ª — REMUNERAÇÃO POR REPASSE PERCENTUAL
7.1. A título de repasse pela prestação dos serviços de docência, a/o ${parteContratada} receberá ${pctProfessor}% (${pctProfessor} por cento) da ${baseCalculoLabel} arrecadada com as turmas objeto deste contrato, cabendo à CONTRATANTE os ${pctEspaco}% (${pctEspaco} por cento) remanescentes.
7.2. Para os fins desta cláusula, ${baseCalculoLabel} corresponde ao valor arrecadado dos alunos pagantes das turmas, deduzidos os tributos incidentes e as taxas de meios de pagamento aplicáveis.${alertaBaseBruta}
7.3. ${faturamentoTexto} A apuração e o repasse ocorrerão com periodicidade ${periodicidade}, mediante emissão do respectivo ${docFiscal} pela/pelo ${parteContratada} sobre o valor do repasse devido, quitado em ${prazoQuitacao}.
7.4. O repasse ora previsto é estritamente variável e proporcional à arrecadação efetiva das turmas, inexistindo qualquer valor fixo garantido, assumindo a/o ${parteContratada} o risco da variação da receita conforme a adesão dos alunos.
7.5. ${retencaoPorTrilha}${compensacaoTexto}`;
  } else if (ehPF) {
    clausula7 = `CLÁUSULA 7ª — REMUNERAÇÃO
7.1. A/O CONTRATADO(A) receberá R$ ${valorHora} (valor Bruto) por hora/aula efetivamente ministrada.
7.2. A remuneração é estritamente variável, não havendo valor fixo garantido nem pagamento de aula não ministrada, ressalvada a hipótese da cláusula 7.6.
7.3. O pagamento será efetuado com periodicidade ${periodicidade}, mediante apuração das horas/aula ministradas e emissão de RPA pela/pelo CONTRATADO(A), quitado em ${prazoQuitacao}.
7.4. Incidem sobre a remuneração as retenções legais aplicáveis ao contribuinte individual: INSS (11%, até o teto), IRRF (conforme tabela progressiva, quando aplicável) e ISS conforme a legislação municipal; a CONTRATANTE recolhe, ainda, a contribuição patronal (20%) incidente sobre a remuneração de contribuintes individuais.
7.5. A/O CONTRATADO(A) obriga-se a manter regular sua condição de contribuinte individual perante o INSS.${compensacaoTexto}`;
  } else if (ehMEI) {
    clausula7 = `CLÁUSULA 7ª — PREÇO E PAGAMENTO
7.1. A CONTRATADA receberá R$ ${valorHora} por hora/aula efetivamente ministrada.
7.2. A remuneração é estritamente variável, não havendo valor fixo garantido nem pagamento de aula não ministrada, ressalvada a hipótese da cláusula 7.6.
7.3. O pagamento será efetuado com periodicidade ${periodicidade}, mediante apuração das horas/aula e emissão de NFS-e pela CONTRATADA, quitado em ${prazoQuitacao}.
7.4. O pagamento é realizado contra nota fiscal, sem retenção previdenciária patronal, por se tratar de MEI em atividade de docência, ressalvada obrigação legal superveniente que venha a determinar o contrário.
7.5. A CONTRATADA é a única responsável pelo recolhimento de seus tributos (DAS-MEI), obrigando-se a manter regular sua condição de microempreendedora individual.${compensacaoTexto}`;
  } else {
    clausula7 = `CLÁUSULA 7ª — PREÇO E PAGAMENTO
7.1. A CONTRATADA receberá R$ ${valorHora} por hora/aula efetivamente ministrada.
7.2. A remuneração é estritamente variável, ressalvada a hipótese da cláusula 7.6.
7.3. O pagamento será efetuado com periodicidade ${periodicidade}, mediante emissão de NFS-e, quitado em ${prazoQuitacao}.
7.4. Incidem sobre a remuneração as retenções tributárias aplicáveis (ISS, IR, CSLL, PIS, COFINS, INSS quando aplicável), conforme o regime tributário da CONTRATADA e a legislação municipal vigente, competindo à CONTRATADA validar tais retenções junto à sua contabilidade.
7.5. A CONTRATADA responde integralmente por seus tributos e pela manutenção de sua regularidade fiscal e societária.${compensacaoTexto}`;
  }

  const clausula8 = ehPF
    ? `CLÁUSULA 8ª — OBRIGAÇÕES DO(A) CONTRATADO(A)
São obrigações da/do CONTRATADO(A): a) prestar os serviços com zelo, diligência e responsabilidade pedagógica; b) emitir RPA correspondente a cada pagamento recebido; c) manter regular sua inscrição como contribuinte individual perante o INSS; d) responder por seus próprios tributos e contribuições; e) zelar pela conservação dos equipamentos e do espaço físico durante o uso; f) observar as normas de segurança do Espaço Liquidificador; g) comunicar à CONTRATANTE, com a antecedência possível, a impossibilidade de ministrar determinada aula.`
    : ehMEI
    ? `CLÁUSULA 8ª — OBRIGAÇÕES DA CONTRATADA
São obrigações da CONTRATADA: a) prestar os serviços com zelo e responsabilidade pedagógica; b) emitir NFS-e correspondente a cada pagamento recebido; c) manter regular sua condição de MEI, incluindo o recolhimento do DAS, a observância do limite de faturamento anual e a compatibilidade da ocupação registrada com a atividade de docência; d) responsabilizar-se integralmente por seus tributos; e) zelar pela conservação dos equipamentos e do espaço; f) observar as normas de segurança do Espaço; g) comunicar à CONTRATANTE, com a antecedência possível, a impossibilidade de ministrar determinada aula.`
    : `CLÁUSULA 8ª — OBRIGAÇÕES DA CONTRATADA
São obrigações da CONTRATADA: a) prestar os serviços com zelo e responsabilidade pedagógica; b) emitir NFS-e correspondente a cada pagamento recebido; c) manter regularidade fiscal e societária da microempresa, bem como CNAE compatível com a atividade de docência; d) responsabilizar-se por seus tributos e pelos atos de seus prepostos; e) zelar pela conservação dos equipamentos e do espaço; f) observar as normas de segurança do Espaço; g) comunicar à CONTRATANTE, com a antecedência possível, a impossibilidade de ministrar determinada aula.`;

  const clausulaProtecaoCrianca = publicoMenor
    ? `\n\nCLÁUSULA 14-A — PROTEÇÃO À CRIANÇA (TURMAS COM ALUNOS MENORES)
Em turmas que incluam alunos menores de idade, a/o ${parteContratada} atua como OPERADORA(OR) dos dados pessoais desses alunos, figurando a CONTRATANTE como CONTROLADORA, tratando tais dados exclusivamente para a finalidade das aulas e sempre no melhor interesse da criança, nos termos do art. 14 da LGPD. É vedado à/ao ${parteContratada} coletar, compartilhar ou divulgar imagens ou dados de alunos menores sem autorização expressa da CONTRATANTE e do responsável legal do aluno. A/O ${parteContratada} observará, ainda, o Estatuto da Criança e do Adolescente (Lei nº 8.069/90), comprometendo-se a não permanecer a sós com aluno menor fora do horário e do ambiente de aula, e a comunicar de imediato à CONTRATANTE qualquer indício de risco ou abuso, nos termos do art. 13 do ECA. Fica facultado à CONTRATANTE exigir, a seu critério, a apresentação de certidão de antecedentes criminais.`
    : '';

  const vigenciaVariante = vigencia === 'Determinada' ? ` O presente contrato vigorará de ${vigenciaInicio} a ${vigenciaFim}.` : '';
  const clausula15 = `CLÁUSULA 15ª — VIGÊNCIA E RESCISÃO
15.1. O presente contrato vigorará por prazo indeterminado a partir de sua assinatura.${vigenciaVariante}
15.2. Qualquer das partes poderá rescindir este contrato imotivadamente, mediante aviso prévio escrito de ${avisoPrevioDias} dias, sem incidência de multa.
15.3. As aulas já agendadas dentro do período de aviso prévio serão normalmente realizadas e remuneradas.
15.4. Não são devidas quaisquer verbas de natureza trabalhista em razão da rescisão, por inexistir vínculo empregatício entre as partes.`;

  const clausulaPrestadorMenor = ehPrestadorMenor
    ? `\n\nCLÁUSULA 15-A — PRESTADOR(A) MENOR DE IDADE
Sendo a/o CONTRATADO(A) menor de 18 (dezoito) anos, o presente contrato é ${idadePrestador >= 16 ? 'assistido' : 'representado'} por seu(sua) representante legal, ${repLegNome}, na condição de ${repLegParentesco}, CPF nº ${repLegCpf}, que assina conjuntamente este instrumento e declara sua ciência e concordância com a integralidade de seus termos.`
    : '';

  const assinaturaContratada = ehPF
    ? `CONTRATADO(A): ${nome}${nomeSocial ? ' ("' + nomeSocial + '")' : ''} — CPF ${cpf}`
    : ehMEI
    ? `CONTRATADA: ${razaoSocial}${nomeSocial ? ' ("' + nomeSocial + '")' : ''} — CNPJ ${cnpj} — Rep.: ${titularNome}`
    : `CONTRATADA: ${razaoSocial}${nomeSocial ? ' ("' + nomeSocial + '")' : ''} — CNPJ ${cnpj} — Rep.: ${repLegal}`;
  const assinaturaRepLegal = ehPrestadorMenor ? `\nRepresentante legal: ${repLegNome} — CPF ${repLegCpf} (${repLegParentesco})` : '';

  const corpo = `${titulo}

QUALIFICAÇÃO DAS PARTES
CONTRATANTE: LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS, empresa individual inscrita no CNPJ nº 28.398.119/0001-83, com sede na Rua Dr. Carvalho de Mendonça, 67, Campos Elíseos, São Paulo/SP, atuando sob o nome operacional Espaço Liquidificador, representada por sua titular Cristiane Socci Leonel, CPF nº 321.132.368-69, doravante CONTRATANTE.

${qualificacao}

As partes regem-se pelos arts. 593 e seguintes do Código Civil${ehMEI ? ', bem como pela Lei Complementar nº 123/2006' : ''}.

${clausula1}

${clausula2}

${clausulas3a5}${clausula6 ? '\n\n' + clausula6 : ''}

${clausula7}

${clausula8}

${clausula9(parteContratada)}

${clausula10(parteContratada)}

${clausula11(parteContratada)}

${clausula12(parteContratada)}

${clausula13(parteContratada, publicoMenor)}

${clausula14(parteContratada)}${clausulaProtecaoCrianca}

${clausula15}${clausulaPrestadorMenor}

${clausula16()}

${clausula17()}

${anexoIAgenda(agenda)}

ASSINATURAS
CONTRATANTE: LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS — CNPJ 28.398.119/0001-83 — Rep.: Cristiane Socci Leonel
${assinaturaContratada}${assinaturaRepLegal}
Testemunhas: 1) __________ 2) __________`;

  return { corpo, ehPrestadorMenor, idadePrestador, bloqueado: false };
}

function montarAnexoIIItens() {
  return [
    'Declaro possuir aptidão para a prática da modalidade de risco, comprovada por Carta de Proficiência assinada por instrutor(a) ou por Termo de Autodeclaração de Aptidão anexo.',
    'Forneço contato de emergência válido para uso em caso de necessidade durante as atividades.',
    'Estou ciente de que a contratação de seguro de acidentes pessoais é recomendada, a título meramente informativo, ficando a adesão a meu exclusivo critério.',
  ];
}

app.post('/contrato-professor/texto', (req, res) => {
  try {
    const resultado = montarTextoContratoProfessor(req.body);
    if (resultado.bloqueado) return res.json({ ok: true, bloqueado: true, motivo: resultado.motivo });
    const ehRisco = req.body.modalidade === 'Circo Aéreo' || req.body.modalidade === 'Acrobática';
    const crypto = require('crypto');
    const versao = crypto.createHash('sha256').update(resultado.corpo).digest('hex').slice(0, 16);
    res.json({
      ok: true, bloqueado: false,
      texto: resultado.corpo,
      ehPrestadorMenor: resultado.ehPrestadorMenor,
      idadePrestador: resultado.idadePrestador,
      anexoIIItens: ehRisco ? montarAnexoIIItens() : [],
      versao,
    });
  } catch (err) {
    console.error('[contrato-professor] erro ao montar texto:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

function gerarTokenAcesso() {
  return require('crypto').randomBytes(16).toString('hex');
}

app.post('/contrato-professor/criar', async (req, res) => {
  const d = req.body || {};
  if (!d.trilha || !d.modalidade || !d.professorTelefone || !d.valorHora) {
    return res.status(400).json({ ok: false, erro: 'Preencha trilha, modalidade, telefone e valor.' });
  }
  if (d.titularidade === 'Professor traz alunos') {
    return res.json({ ok: true, bloqueado: true, motivo: 'Neste modelo, o(a) professor(a) utiliza o espaço mediante aluguel (cessão onerosa de uso), e não um contrato de docência. Utilize o instrumento de Cessão Onerosa de Uso para este caso.' });
  }
  try {
    const idadePrestador = calcularIdade(d.dataNascimento);
    const ehPrestadorMenor = idadePrestador !== null && idadePrestador < 18;
    if (ehPrestadorMenor && (!d.repLegNome || !d.repLegCpf)) {
      return res.status(400).json({ ok: false, erro: 'Prestador menor de 18 anos exige nome e CPF do representante legal.' });
    }
    if (d.titularidade === 'Misto') {
      const soma = Number(d.pctProfessor || 0) + Number(d.pctEspaco || 0);
      if (soma !== 100) return res.status(400).json({ ok: false, erro: 'A soma dos percentuais (professor + espaço) deve ser 100%.' });
    }

    const token = gerarTokenAcesso();
    const nomeExibicao = d.nomeSocial || (d.trilha === 'PF-RPA' ? (d.nome || '') : (d.razaoSocial || ''));
    const riscoAlto = d.trilha === 'PF-RPA' && d.titularidade === 'Espaço matricula/cobra' && d.formato === 'Regulares contínuas';
    const cpfCnpj = (d.trilha === 'PF-RPA' ? d.cpf : d.cnpj) || '';

    const props = {
      'Título': { title: [{ text: { content: nomeExibicao + ' — ' + d.modalidade } }] },
      'Professor': { rich_text: [{ text: { content: nomeExibicao } }] },
      'CPF/CNPJ': { rich_text: [{ text: { content: cpfCnpj } }] },
      'Trilha': { select: { name: d.trilha } },
      'Modalidade': { select: { name: d.modalidade } },
      'Titularidade dos Alunos': { select: { name: d.titularidade } },
      'Formato': { select: { name: d.formato } },
      'Valor Hora/Aula': { number: Number(d.valorHora) },
      'Status': { select: { name: 'Aguardando Aceite' } },
      'Token de Acesso': { rich_text: [{ text: { content: token } }] },
      'Alerta Risco Alto': { checkbox: !!riscoAlto },
      'Dados JSON': { rich_text: [{ text: { content: JSON.stringify(d).slice(0, 1900) } }] },
    };

    const rCriar = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: CONTRATOS_PROF_DB }, properties: props }),
    });
    if (!rCriar.ok) { const t = await rCriar.text(); throw new Error('Notion: ' + t); }

    const numLimpo = (d.professorTelefone || '').replace(/\D/g, '');
    const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
    const linkContrato = 'https://contratos-professores.ciadoliquidificador.com.br/aceitar/?token=' + token;
    if (numBr) {
      try {
        await enviarWhatsApp(numBr, 'Olá, ' + (nomeExibicao.split(' ')[0] || '') + '! 📄\n\nSeu contrato de prestação de serviços de docência (' + d.modalidade + ') está pronto para revisão e aceite:\n\n' + linkContrato + '\n\nQualquer dúvida, é só chamar.');
      } catch(e) {}
    }
    if (riscoAlto) {
      try { await enviarWhatsApp(WHATSAPP_FABIO, '⚠️ *Alerta de risco alto* — contrato de ' + nomeExibicao + ' (PF, Espaço matricula/cobra, aulas contínuas). Considere a trilha PJ.'); } catch(e) {}
    }

    res.json({ ok: true, bloqueado: false, token, link: linkContrato, alertaRiscoAlto: riscoAlto });
  } catch (err) {
    console.error('[contrato-professor/criar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.get('/contrato-professor/rascunho/:token', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROF_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Token de Acesso', rich_text: { equals: req.params.token } }, page_size: 1 }),
    });
    const dResp = await r.json();
    const pagina = (dResp.results || [])[0];
    if (!pagina) return res.json({ ok: false, erro: 'Contrato não encontrado.' });
    const p = pagina.properties;
    const jsonBruto = p['Dados JSON']?.rich_text?.[0]?.plain_text || '{}';
    let dadosCompletos = {};
    try { dadosCompletos = JSON.parse(jsonBruto); } catch(e) {}
    res.json({ ok: true, pageId: pagina.id, status: p['Status']?.select?.name, dadosCompletos });
  } catch (err) {
    console.error('[contrato-professor/rascunho] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/contrato-professor/aceitar', async (req, res) => {
  const { token, assinaturaDigitada, anexoIIAceites, consentimentoLgpd, consentimentoRepLegal, versaoContrato, dispositivo } = req.body;
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

  if (!token || !assinaturaDigitada || !consentimentoLgpd) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROF_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Token de Acesso', rich_text: { equals: token } }, page_size: 1 }),
    });
    const dResp = await r.json();
    const pagina = (dResp.results || [])[0];
    if (!pagina) return res.status(404).json({ ok: false, erro: 'Contrato não encontrado.' });

    const jsonBruto = pagina.properties['Dados JSON']?.rich_text?.[0]?.plain_text || '{}';
    const dadosCompletos = JSON.parse(jsonBruto);

    const resultado = montarTextoContratoProfessor(dadosCompletos);
    if (resultado.bloqueado) return res.status(400).json({ ok: false, erro: resultado.motivo });
    if (resultado.ehPrestadorMenor && !consentimentoRepLegal) {
      return res.status(400).json({ ok: false, erro: 'É necessário o consentimento do representante legal.' });
    }
    const ehRisco = dadosCompletos.modalidade === 'Circo Aéreo' || dadosCompletos.modalidade === 'Acrobática';
    if (ehRisco && (!Array.isArray(anexoIIAceites) || anexoIIAceites.some(v => !v))) {
      return res.status(400).json({ ok: false, erro: 'Aceite todos os itens do Anexo II (Aptidão e Segurança).' });
    }

    const dataHoraISO = new Date().toISOString();
    const crypto = require('crypto');
    const versaoCalculada = versaoContrato || crypto.createHash('sha256').update(resultado.corpo).digest('hex').slice(0, 16);

    let linkPdf = '';
    try {
      const PDFDocument = require('pdfkit');
      const chunks = [];
      const doc = new PDFDocument({ margin: 50 });
      doc.on('data', c => chunks.push(c));
      const fim = new Promise(resolve => doc.on('end', () => resolve(Buffer.concat(chunks))));
      doc.fontSize(9).text(resultado.corpo);
      doc.moveDown(2);
      doc.fontSize(9).text('--- REGISTRO DE ACEITE ELETRÔNICO ---');
      doc.text('Assinado por: ' + assinaturaDigitada);
      doc.text('Data/Hora: ' + dataHoraISO);
      doc.text('IP: ' + ip);
      doc.text('Dispositivo: ' + (dispositivo || ''));
      doc.text('Versão do documento: ' + versaoCalculada);
      doc.end();
      const pdfBuffer = await fim;

      const msToken = await getMicrosoftToken();
      const nomePasta = 'Contratos-Professores-' + slugify(dadosCompletos.trilha) + '-' + slugify(assinaturaDigitada);
      const folderId = await criarOuObterSubpasta(msToken, nomePasta);
      const pdfUploaded = await uploadBufferOneDrive(msToken, folderId, pdfBuffer, 'contrato-' + Date.now() + '.pdf');
      linkPdf = await criarLinkCompartilhamento(msToken, pdfUploaded.id);
    } catch (e) {
      console.error('[contrato-professor/aceitar] erro ao gerar/subir PDF:', e.message);
    }

    await fetch('https://api.notion.com/v1/pages/' + pagina.id, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Status': { select: { name: 'Aceito' } },
          'Assinatura Digitada': { rich_text: [{ text: { content: assinaturaDigitada } }] },
          'IP Aceite': { rich_text: [{ text: { content: ip } }] },
          'User-Agent Aceite': { rich_text: [{ text: { content: dispositivo || '' } }] },
          'Data/Hora Aceite': { date: { start: dataHoraISO } },
          'Versão do Contrato': { rich_text: [{ text: { content: versaoCalculada } }] },
          'Link do Contrato PDF': { url: linkPdf || null },
        },
      }),
    });

    try { await enviarWhatsApp(WHATSAPP_FABIO, '✅ *Contrato aceito* — ' + assinaturaDigitada + ' (' + dadosCompletos.modalidade + ', ' + dadosCompletos.trilha + ')'); } catch(e) {}

    // Auto-cadastro no Portal Profs (banco "Professores — Cadastro"), evitando
    // digitacao duplicada manual. So cria se ainda nao existir um cadastro com
    // esse CPF; nunca sobrescreve um cadastro ja existente.
    try {
      const cpfParaCadastro = (
        dadosCompletos.trilha === 'PF-RPA' ? dadosCompletos.cpf :
        dadosCompletos.trilha === 'PJ-MEI' ? dadosCompletos.titularCpf :
        dadosCompletos.trilha === 'PJ-ME' ? dadosCompletos.repCpf : ''
      ) || '';
      const cpfLimpoCadastro = cpfParaCadastro.replace(/\D/g, '');

      if (cpfLimpoCadastro) {
        const jaCadastrado = await buscarProfessorPorCpf(cpfLimpoCadastro);
        if (!jaCadastrado) {
          const rCriarCadastro = await fetch('https://api.notion.com/v1/pages', {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
            body: JSON.stringify({
              parent: { database_id: PROFESSORES_CADASTRO_DB },
              properties: {
                'Nome': { title: [{ text: { content: assinaturaDigitada } }] },
                'CPF': { rich_text: [{ text: { content: cpfLimpoCadastro } }] },
                'Telefone': { phone_number: dadosCompletos.professorTelefone || null },
                'Endereço': { rich_text: [{ text: { content: dadosCompletos.endereco || '' } }] },
              },
            }),
          });
          if (rCriarCadastro.ok) {
            try {
              await enviarWhatsApp(WHATSAPP_FABIO, '👤 Cadastro criado automaticamente no Portal Profs para ' + assinaturaDigitada + ' — já pode fazer login.');
            } catch (eNotifCad) {}
          } else {
            const eBodyCad = await rCriarCadastro.text();
            console.error('[contrato-professor/aceitar] Notion recusou criar cadastro automático:', eBodyCad);
          }
        }
      } else {
        console.error('[contrato-professor/aceitar] sem CPF disponível para auto-cadastro (trilha: ' + dadosCompletos.trilha + ').');
      }
    } catch (eAutoCad) {
      console.error('[contrato-professor/aceitar] erro no auto-cadastro:', eAutoCad.message);
    }

    res.json({ ok: true, linkPdf });
  } catch (err) {
    console.error('[contrato-professor/aceitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

const ADITIVOS_DB = '9b35319a73854e318cd9efc6497bfb0e';

const CAMPOS_ADITAVEIS_LABELS = {
  valorHora: 'Cláusula 7.1 (valor da hora/aula)',
  agenda: 'Anexo I (agenda)',
  pctProfessor: 'Cláusula 7.1 (percentual do(a) professor(a))',
  pctEspaco: 'Cláusula 7.1 (percentual do Espaço)',
  baseCalculo: 'Cláusula 7.2 (base de cálculo do repasse)',
  quemFatura: 'Cláusula 7.3 (quem fatura o aluno pagante)',
  periodicidade: 'Cláusula 7.3 (periodicidade de pagamento)',
  prazoQuitacao: 'Cláusula 7.3 (prazo de quitação)',
  avisoPrevioDias: 'Cláusula 15.2 (aviso prévio de rescisão)',
  compensacaoCancelamento: 'Cláusula 7.6 (compensação por cancelamento)',
  compensacaoHoras: 'Cláusula 7.6 (horas de antecedência da compensação)',
  compensacaoPct: 'Cláusula 7.6 (percentual da compensação)',
  vigencia: 'Cláusula 15.1 (vigência)',
  vigenciaInicio: 'Cláusula 15.1 (início da vigência)',
  vigenciaFim: 'Cláusula 15.1 (fim da vigência)',
  substituicao: 'Cláusula 6ª (faculdade de substituição)',
  formato: 'Cláusula 1.1 (formato da oferta)',
  modalidade: 'Cláusula 1.1 (modalidade) — atenção: pode alterar anexos de aptidão ou proteção à criança',
  publicoMenor: 'Cláusula 14-A (proteção à criança)',
};

const CAMPOS_BLOQUEADOS_ADITIVO = ['trilha', 'titularidade'];

async function buscarAditivosDoContrato(contratoId) {
  const r = await fetch('https://api.notion.com/v1/databases/' + ADITIVOS_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { and: [
        { property: 'Contrato Original', relation: { contains: contratoId } },
        { property: 'Status', select: { equals: 'Aceito' } },
      ]},
      page_size: 100,
    }),
  });
  const d = await r.json();
  const aditivos = (d.results || []).map(p => {
    const props = p.properties;
    let alteracoes = [];
    try { alteracoes = JSON.parse(props['Alterações JSON']?.rich_text?.[0]?.plain_text || '[]'); } catch(e) {}
    return { numero: props['Número']?.number || 0, alteracoes };
  });
  aditivos.sort((a, b) => a.numero - b.numero);
  return aditivos;
}

function aplicarAlteracoes(estadoBase, alteracoes) {
  const novoEstado = { ...estadoBase };
  alteracoes.forEach(alt => { novoEstado[alt.campo] = alt.valorNovo; });
  return novoEstado;
}

app.get('/contrato-professor/listar-assinados', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROF_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Status', select: { equals: 'Aceito' } }, page_size: 100 }),
    });
    const d = await r.json();
    const contratos = (d.results || []).map(p => ({
      pageId: p.id,
      professor: p.properties['Professor']?.rich_text?.[0]?.plain_text || '',
      modalidade: p.properties['Modalidade']?.select?.name || '',
      trilha: p.properties['Trilha']?.select?.name || '',
    }));
    res.json({ ok: true, contratos });
  } catch (err) {
    console.error('[aditivo/listar-assinados] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.get('/contrato-professor/estado-vigente/:pageId', async (req, res) => {
  try {
    const rOriginal = await fetch('https://api.notion.com/v1/pages/' + req.params.pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const original = await rOriginal.json();
    if (!original.properties) return res.json({ ok: false, erro: 'Contrato não encontrado.' });
    if (original.properties['Status']?.select?.name !== 'Aceito') {
      return res.json({ ok: false, erro: 'Só é possível aditar contratos já assinados.' });
    }

    const dadosOriginais = JSON.parse(original.properties['Dados JSON']?.rich_text?.[0]?.plain_text || '{}');
    const aditivosAnteriores = await buscarAditivosDoContrato(req.params.pageId);
    let estadoVigente = { ...dadosOriginais };
    aditivosAnteriores.forEach(ad => { estadoVigente = aplicarAlteracoes(estadoVigente, ad.alteracoes); });

    res.json({
      ok: true,
      pageId: req.params.pageId,
      professor: original.properties['Professor']?.rich_text?.[0]?.plain_text || '',
      cpfCnpj: original.properties['CPF/CNPJ']?.rich_text?.[0]?.plain_text || '',
      proximoNumero: aditivosAnteriores.length + 1,
      estadoVigente,
    });
  } catch (err) {
    console.error('[aditivo/estado-vigente] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/contrato-professor/aditivo/criar', async (req, res) => {
  const { contratoOriginalId, campos, dataEfeito } = req.body;
  if (!contratoOriginalId || !campos || !dataEfeito) {
    return res.status(400).json({ ok: false, erro: 'Preencha os campos obrigatórios.' });
  }
  const camposBloqueadosEnviados = Object.keys(campos).filter(c => CAMPOS_BLOQUEADOS_ADITIVO.includes(c));
  if (camposBloqueadosEnviados.length > 0) {
    return res.status(400).json({ ok: false, erro: 'Mudança de tipo de prestador (PF/MEI/ME) ou de modelo de remuneração (titularidade) exige um contrato NOVO, não um aditivo.' });
  }

  try {
    const rOriginal = await fetch('https://api.notion.com/v1/pages/' + contratoOriginalId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const original = await rOriginal.json();
    if (!original.properties) return res.status(404).json({ ok: false, erro: 'Contrato não encontrado.' });
    if (original.properties['Status']?.select?.name !== 'Aceito') {
      return res.status(400).json({ ok: false, erro: 'Só é possível aditar contratos já assinados.' });
    }

    const dadosOriginais = JSON.parse(original.properties['Dados JSON']?.rich_text?.[0]?.plain_text || '{}');
    const aditivosAnteriores = await buscarAditivosDoContrato(contratoOriginalId);
    let estadoVigente = { ...dadosOriginais };
    aditivosAnteriores.forEach(ad => { estadoVigente = aplicarAlteracoes(estadoVigente, ad.alteracoes); });

    if ('pctProfessor' in campos || 'pctEspaco' in campos) {
      const novoPctProf = Number('pctProfessor' in campos ? campos.pctProfessor : estadoVigente.pctProfessor);
      const novoPctEsp = Number('pctEspaco' in campos ? campos.pctEspaco : estadoVigente.pctEspaco);
      if (novoPctProf + novoPctEsp !== 100) {
        return res.status(400).json({ ok: false, erro: 'A soma dos percentuais deve ser 100%.' });
      }
    }

    const alteracoes = Object.entries(campos)
      .filter(([campo, novoValor]) => String(estadoVigente[campo] ?? '') !== String(novoValor))
      .map(([campo, novoValor]) => ({
        campo,
        referenciaClausula: CAMPOS_ADITAVEIS_LABELS[campo] || campo,
        valorAnterior: estadoVigente[campo] ?? '(não definido)',
        valorNovo: novoValor,
      }));

    if (alteracoes.length === 0) {
      return res.status(400).json({ ok: false, erro: 'Nenhuma alteração real foi detectada em relação ao estado vigente.' });
    }

    const numero = aditivosAnteriores.length + 1;
    const token = require('crypto').randomBytes(16).toString('hex');
    const professor = original.properties['Professor']?.rich_text?.[0]?.plain_text || '';
    const cpfCnpj = original.properties['CPF/CNPJ']?.rich_text?.[0]?.plain_text || '';

    await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: ADITIVOS_DB },
        properties: {
          'Título': { title: [{ text: { content: professor + ' — ' + numero + 'º Aditivo' } }] },
          'Contrato Original': { relation: [{ id: contratoOriginalId }] },
          'CPF/CNPJ Contrato': { rich_text: [{ text: { content: cpfCnpj } }] },
          'Número': { number: numero },
          'Data de Efeito': { date: { start: dataEfeito } },
          'Alterações JSON': { rich_text: [{ text: { content: JSON.stringify(alteracoes).slice(0, 1900) } }] },
          'Status': { select: { name: 'Aguardando Aceite' } },
          'Token de Acesso': { rich_text: [{ text: { content: token } }] },
        },
      }),
    });

    const telefone = (dadosOriginais.professorTelefone || '').replace(/\D/g, '');
    const numBr = telefone.length === 11 ? '55' + telefone : telefone;
    const linkAditivo = 'https://contratos-professores.ciadoliquidificador.com.br/aditivo/?token=' + token;
    if (numBr) {
      try {
        await enviarWhatsApp(numBr, 'Olá, ' + (professor.split(' ')[0] || '') + '! 📎\n\nHá um Termo Aditivo (nº ' + numero + ') ao seu contrato de docência aguardando sua revisão e aceite:\n\n' + linkAditivo + '\n\nQualquer dúvida, é só chamar.');
      } catch(e) {}
    }

    res.json({ ok: true, numero, token, link: linkAditivo, alteracoes });
  } catch (err) {
    console.error('[aditivo/criar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

function montarTextoTermoAditivo(dados) {
  const { professor, numero, dataOriginalFmt, alteracoes, dataEfeitoFmt, dataAssinaturaFmt, contratoOriginalId } = dados;
  const listaAlteracoes = alteracoes.map(a =>
    '• ' + a.referenciaClausula + ' — De: ' + a.valorAnterior + '  →  Para: ' + a.valorNovo
  ).join('\n');

  return `TERMO ADITIVO Nº ${numero} AO CONTRATO DE PRESTAÇÃO DE SERVIÇOS DE DOCÊNCIA

Referência: contrato firmado em ${dataOriginalFmt} (ID ${contratoOriginalId}).

PARTES: CONTRATANTE — LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS, CNPJ nº 28.398.119/0001-83, representada por sua titular Cristiane Socci Leonel, CPF nº 321.132.368-69; e CONTRATADO(A) — ${professor}, conforme qualificação constante do contrato original acima referenciado.

Considerando que as partes desejam, de comum acordo, alterar termos do contrato acima identificado, firmam o presente Termo Aditivo:

CLÁUSULA 1ª — DAS ALTERAÇÕES
Ficam alteradas as seguintes disposições:
${listaAlteracoes}

CLÁUSULA 2ª — DA VIGÊNCIA DAS ALTERAÇÕES
2.1. As alterações ora pactuadas produzem efeitos a partir de ${dataEfeitoFmt}.

CLÁUSULA 3ª — DA RATIFICAÇÃO
3.1. Ficam ratificadas e inalteradas todas as demais cláusulas e condições do contrato original, que permanece em pleno vigor no que não conflitar com o presente Termo Aditivo.

São Paulo, ${dataAssinaturaFmt}.

ASSINATURAS
CONTRATANTE: LIQUIDIFICADOR PRODUÇÕES ARTÍSTICAS — Rep.: Cristiane Socci Leonel
CONTRATADO(A): ${professor}
Testemunhas: 1) __________ 2) __________`;
}

app.get('/contrato-professor/aditivo/rascunho/:token', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ADITIVOS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Token de Acesso', rich_text: { equals: req.params.token } }, page_size: 1 }),
    });
    const dResp = await r.json();
    const pagina = (dResp.results || [])[0];
    if (!pagina) return res.json({ ok: false, erro: 'Termo aditivo não encontrado.' });

    const contratoRel = pagina.properties['Contrato Original']?.relation?.[0]?.id;
    const rOriginal = await fetch('https://api.notion.com/v1/pages/' + contratoRel, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const original = await rOriginal.json();
    const professor = original.properties['Professor']?.rich_text?.[0]?.plain_text || '';
    const dataOriginalISO = original.created_time ? original.created_time.split('T')[0] : '';

    let alteracoes = [];
    try { alteracoes = JSON.parse(pagina.properties['Alterações JSON']?.rich_text?.[0]?.plain_text || '[]'); } catch(e) {}
    const numero = pagina.properties['Número']?.number || 0;
    const dataEfeitoISO = pagina.properties['Data de Efeito']?.date?.start || '';

    const texto = montarTextoTermoAditivo({
      professor, numero,
      dataOriginalFmt: dataOriginalISO ? dataOriginalISO.split('-').reverse().join('/') : '(data não disponível)',
      alteracoes,
      dataEfeitoFmt: dataEfeitoISO ? dataEfeitoISO.split('-').reverse().join('/') : '(a definir)',
      dataAssinaturaFmt: new Date().toLocaleDateString('pt-BR'),
      contratoOriginalId: contratoRel,
    });

    res.json({ ok: true, pageId: pagina.id, professor, numero, alteracoes, texto });
  } catch (err) {
    console.error('[aditivo/rascunho] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/contrato-professor/aditivo/aceitar', async (req, res) => {
  const { token, assinaturaDigitada, consentimentoLgpd, dispositivo } = req.body;
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

  if (!token || !assinaturaDigitada || !consentimentoLgpd) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ADITIVOS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Token de Acesso', rich_text: { equals: token } }, page_size: 1 }),
    });
    const dResp = await r.json();
    const pagina = (dResp.results || [])[0];
    if (!pagina) return res.status(404).json({ ok: false, erro: 'Termo aditivo não encontrado.' });

    const contratoRel = pagina.properties['Contrato Original']?.relation?.[0]?.id;
    const rOriginal = await fetch('https://api.notion.com/v1/pages/' + contratoRel, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const original = await rOriginal.json();
    const professor = original.properties['Professor']?.rich_text?.[0]?.plain_text || '';
    const dataOriginalISO = original.created_time ? original.created_time.split('T')[0] : '';

    let alteracoes = [];
    try { alteracoes = JSON.parse(pagina.properties['Alterações JSON']?.rich_text?.[0]?.plain_text || '[]'); } catch(e) {}
    const numero = pagina.properties['Número']?.number || 0;
    const dataEfeitoISO = pagina.properties['Data de Efeito']?.date?.start || '';

    const dataHoraISO = new Date().toISOString();
    const texto = montarTextoTermoAditivo({
      professor, numero,
      dataOriginalFmt: dataOriginalISO ? dataOriginalISO.split('-').reverse().join('/') : '(data não disponível)',
      alteracoes,
      dataEfeitoFmt: dataEfeitoISO ? dataEfeitoISO.split('-').reverse().join('/') : '(a definir)',
      dataAssinaturaFmt: new Date().toLocaleDateString('pt-BR'),
      contratoOriginalId: contratoRel,
    });

    const crypto = require('crypto');
    const versaoCalculada = crypto.createHash('sha256').update(texto).digest('hex').slice(0, 16);

    let linkPdf = '';
    try {
      const PDFDocument = require('pdfkit');
      const chunks = [];
      const doc = new PDFDocument({ margin: 50 });
      doc.on('data', c => chunks.push(c));
      const fim = new Promise(resolve => doc.on('end', () => resolve(Buffer.concat(chunks))));
      doc.fontSize(9).text(texto);
      doc.moveDown(2);
      doc.fontSize(9).text('--- REGISTRO DE ACEITE ELETRÔNICO ---');
      doc.text('Assinado por: ' + assinaturaDigitada);
      doc.text('Data/Hora: ' + dataHoraISO);
      doc.text('IP: ' + ip);
      doc.text('Dispositivo: ' + (dispositivo || ''));
      doc.text('Versão do documento: ' + versaoCalculada);
      doc.end();
      const pdfBuffer = await fim;

      const msToken = await getMicrosoftToken();
      const nomePasta = 'Aditivos-Contratos-Professores-' + slugify(professor);
      const folderId = await criarOuObterSubpasta(msToken, nomePasta);
      const pdfUploaded = await uploadBufferOneDrive(msToken, folderId, pdfBuffer, 'aditivo-' + numero + '-' + Date.now() + '.pdf');
      linkPdf = await criarLinkCompartilhamento(msToken, pdfUploaded.id);
    } catch (e) {
      console.error('[aditivo/aceitar] erro ao gerar/subir PDF:', e.message);
    }

    await fetch('https://api.notion.com/v1/pages/' + pagina.id, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Status': { select: { name: 'Aceito' } },
          'Assinatura Digitada': { rich_text: [{ text: { content: assinaturaDigitada } }] },
          'IP Aceite': { rich_text: [{ text: { content: ip } }] },
          'User-Agent Aceite': { rich_text: [{ text: { content: dispositivo || '' } }] },
          'Data/Hora Aceite': { date: { start: dataHoraISO } },
          'Versão do Documento': { rich_text: [{ text: { content: versaoCalculada } }] },
          'Link do PDF': { url: linkPdf || null },
        },
      }),
    });

    try { await enviarWhatsApp(WHATSAPP_FABIO, '✅ *Termo Aditivo aceito* — ' + professor + ' (nº ' + numero + ')'); } catch(e) {}

    res.json({ ok: true, linkPdf });
  } catch (err) {
    console.error('[aditivo/aceitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.get('/contrato-professor/meu/:cpfCnpj', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + CONTRATOS_PROF_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'CPF/CNPJ', rich_text: { equals: req.params.cpfCnpj } }, page_size: 20 }),
    });
    const dResp = await r.json();
    const contratos = (dResp.results || []).map(pagina => {
      const p = pagina.properties;
      return {
        pageId: pagina.id,
        modalidade: p['Modalidade']?.select?.name,
        trilha: p['Trilha']?.select?.name,
        status: p['Status']?.select?.name,
        linkPdf: p['Link do Contrato PDF']?.url || '',
        dataAceite: p['Data/Hora Aceite']?.date?.start || '',
        valorHora: p['Valor Hora/Aula']?.number || null,
        periodicidade: p['Periodicidade']?.select?.name || '',
        vigencia: p['Vigência']?.select?.name || '',
        solicitacaoPendente: p['Solicitação Pendente']?.select?.name || 'Nenhuma',
      };
    });
    res.json({ ok: true, contratos });
  } catch (err) {
    console.error('[contrato-professor/meu] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/contrato-professor/solicitar-cancelamento', async (req, res) => {
  const { pageId, nome, motivo } = req.body;
  if (!pageId || !nome) return res.status(400).json({ ok: false, erro: 'Dados incompletos.' });
  try {
    await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Solicitação Pendente': { select: { name: 'Cancelamento' } },
          'Detalhes da Solicitação': { rich_text: [{ text: { content: motivo || '(sem motivo informado)' } }] },
        },
      }),
    });
    try { await enviarWhatsApp(WHATSAPP_FABIO, '⚠️ *Professor solicitou cancelamento de contrato* — ' + nome + '\nMotivo: ' + (motivo || '(sem motivo informado)')); } catch(e) {}
    res.json({ ok: true });
  } catch (err) {
    console.error('[contrato-professor/solicitar-cancelamento] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/contrato-professor/propor-mudanca', async (req, res) => {
  const { pageId, nome, proposta } = req.body;
  if (!pageId || !nome || !proposta) return res.status(400).json({ ok: false, erro: 'Descreva a mudança desejada.' });
  try {
    await fetch('https://api.notion.com/v1/pages/' + pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Solicitação Pendente': { select: { name: 'Mudança' } },
          'Detalhes da Solicitação': { rich_text: [{ text: { content: proposta } }] },
        },
      }),
    });
    try { await enviarWhatsApp(WHATSAPP_FABIO, '📝 *Professor propôs mudança no contrato* — ' + nome + '\nProposta: ' + proposta); } catch(e) {}
    res.json({ ok: true });
  } catch (err) {
    console.error('[contrato-professor/propor-mudanca] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// PORTAL ALUNA — login por CPF, contrato, presenca, solicitacoes
// ============================================================
const PRESENCAS_2026_DB = '282c0bc0-06d8-4828-acae-d5ac35388318';
const REPOSICOES_DB = 'dde8519e6e0f4157b2bb56b545e2ef84';

// ============================================================
// COTA DE REPOSIÇÃO — créditos individuais com janela rolante de 30 dias
// Cada Falta gera um crédito (Prazo Limite = Data da falta + 30 dias).
// A cota é quantos créditos "Aberto" simultâneos a aluna pode ter por modalidade.
// ============================================================
function calcularCotaReposicao(frequencia) {
  if (frequencia === '1x semana') return 1;
  if (frequencia === '2x semana') return 2;
  return 1; // fallback seguro para "Acordo" ou qualquer valor inesperado
}

function somarDias(dataISO, dias) {
  const d = new Date(dataISO + 'T12:00:00Z'); // meio-dia UTC evita virada de dia por fuso horário
  d.setUTCDate(d.getUTCDate() + dias);
  return d.toISOString().slice(0, 10);
}

async function buscarFrequenciaAluna(cpfLimpo, modalidade) {
  const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { and: [
        { property: 'CPF', rich_text: { equals: cpfLimpo } },
        { property: 'Modalidade', select: { equals: modalidade } },
        { property: 'Status', select: { equals: 'Ativa' } },
      ]},
      page_size: 5,
    }),
  });
  const d = await r.json();
  const pagina = (d.results || [])[0];
  return pagina?.properties?.['Frequência']?.select?.name || '';
}

// Retorna { cota, creditosAbertos, podeAgendar, creditos: [{id, dataFalta, prazoLimite, turmaOrigem}] }
// Expira preguiçosamente (lazy) qualquer crédito "Aberto" cujo Prazo Limite já passou.
async function verificarCotaReposicao(cpfLimpo, modalidade) {
  const frequencia = await buscarFrequenciaAluna(cpfLimpo, modalidade);
  const cota = calcularCotaReposicao(frequencia);

  const r = await fetch('https://api.notion.com/v1/databases/' + REPOSICOES_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { and: [
        { property: 'CPF Aluna', rich_text: { equals: cpfLimpo } },
        { property: 'Modalidade', rich_text: { equals: modalidade } },
        { property: 'Status', select: { equals: 'Aberto' } },
      ]},
      page_size: 100,
    }),
  });
  const d = await r.json();
  const hoje = new Date().toISOString().slice(0, 10);

  const creditos = [];
  for (const pagina of (d.results || [])) {
    const prazoLimite = pagina.properties?.['Prazo Limite']?.date?.start || '';
    if (prazoLimite && prazoLimite < hoje) {
      try {
        await fetch('https://api.notion.com/v1/pages/' + pagina.id, {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
          body: JSON.stringify({ properties: { 'Status': { select: { name: 'Expirado' } } } }),
        });
      } catch (e) { console.error('[reposicao] erro ao expirar credito:', e.message); }
      continue; // expirado, nao conta pra cota
    }
    creditos.push({
      id: pagina.id,
      dataFalta: prazoLimite ? somarDias(prazoLimite, -30) : '',
      prazoLimite,
      turmaOrigem: pagina.properties?.['Turma Origem']?.rich_text?.[0]?.plain_text || '',
    });
  }
  creditos.sort((a, b) => a.prazoLimite.localeCompare(b.prazoLimite)); // mais antigo (prazo mais proximo) primeiro

  // "No limite mas nao bloqueada": com creditos.length === cota, ela ainda pode marcar (consome o mais
  // antigo). So bloqueia quando ja excede a cota (creditos acumulados sem serem usados) ou nao ha credito algum.
  const podeAgendar = creditos.length > 0 && creditos.length <= cota;
  const mensagemBloqueio = creditos.length === 0
    ? 'Não encontramos nenhuma falta em aberto para repor. Se você acredita que isso é um engano, fale com a gente.'
    : 'Você já tem uma reposição em aberto. Assim que ela for utilizada ou expirar, você poderá agendar a próxima.';
  return { cota, creditosAbertos: creditos.length, podeAgendar, creditos, mensagemBloqueio };
}

// Vincula a marcação da reposição ao crédito e muda Status para "Usado"
async function consumirCreditoReposicao(creditoId, { turmaDesejada, dataDesejada }) {
  await fetch('https://api.notion.com/v1/pages/' + creditoId, {
    method: 'PATCH',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      properties: {
        'Data Desejada': { rich_text: [{ text: { content: dataDesejada || '' } }] },
        'Turma Desejada': { rich_text: [{ text: { content: turmaDesejada || '' } }] },
        'Status': { select: { name: 'Usado' } },
      },
    }),
  });
}

// Chamado logo apos gravar uma Falta em Presenças 2026: cria o credito de reposicao correspondente.
// Idempotente: se ja existir um credito com essa "Falta Relacionada", nao duplica.
async function criarCreditoReposicaoPorFalta({ faltaPageId, alunaId, nomeAluna, modalidade, turma, dataFalta }) {
  try {
    const rExiste = await fetch('https://api.notion.com/v1/databases/' + REPOSICOES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { property: 'Falta Relacionada', relation: { contains: faltaPageId } },
        page_size: 1,
      }),
    });
    const dExiste = await rExiste.json();
    if ((dExiste.results || []).length > 0) return; // credito ja existe para essa falta

    const alunaResp = await fetch('https://api.notion.com/v1/pages/' + alunaId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    if (!alunaResp.ok) { console.error('[reposicao] aluna nao encontrada para credito:', alunaId); return; }
    const alunaData = await alunaResp.json();
    const cpf = (alunaData.properties?.CPF?.rich_text?.[0]?.plain_text || '').replace(/\D/g, '');
    if (!cpf) { console.error('[reposicao] aluna sem CPF, credito nao criado:', alunaId); return; }

    const prazoLimite = somarDias(dataFalta, 30);

    await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: REPOSICOES_DB },
        properties: {
          'Título': { title: [{ text: { content: `${nomeAluna} - Falta ${dataFalta} - ${modalidade}` } }] },
          'CPF Aluna': { rich_text: [{ text: { content: cpf } }] },
          'Modalidade': { rich_text: [{ text: { content: modalidade } }] },
          'Turma Origem': { rich_text: [{ text: { content: turma || '' } }] },
          'Falta Relacionada': { relation: [{ id: faltaPageId }] },
          'Prazo Limite': { date: { start: prazoLimite } },
          'Status': { select: { name: 'Aberto' } },
        },
      }),
    });
  } catch (e) {
    console.error('[reposicao] erro ao criar credito por falta:', e.message);
  }
}

async function buscarMatriculasPorCpf(cpfLimpo) {
  const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      filter: { and: [
        { property: 'CPF', rich_text: { equals: cpfLimpo } },
        { property: 'Status', select: { equals: 'Ativa' } },
      ]},
      page_size: 20,
    }),
  });
  const d = await r.json();
  return (d.results || []).map(p => {
    const props = p.properties;
    return {
      pageId: p.id,
      nome: props['Nome']?.title?.[0]?.plain_text || '',
      contato: props['Contato']?.phone_number || '',
      email: props['Email']?.email || '',
      endereco: props['Endereço']?.rich_text?.[0]?.plain_text || '',
      modalidade: props['Modalidade']?.select?.name || '',
      turma: props['Turma']?.select?.name || '',
      dia: props['Dia']?.select?.name || '',
      horario: props['Horário']?.select?.name || '',
      professor: props['Professor']?.select?.name || '',
      plano: props['Plano']?.select?.name || '',
      frequencia: props['Frequência']?.select?.name || '',
      valor: props['Valor']?.number || null,
      linkContratoPdf: props['Link do Contrato PDF']?.url || '',
      contatoEmergenciaNome: props['Contato de Emergência']?.rich_text?.[0]?.plain_text || '',
      contatoEmergenciaTelefone: props['Tel. Emergência']?.phone_number || '',
      // Data de inicio do plano em curso = data em que o contrato foi aceito (Clausula Primeira: "Validade: a partir de ...").
      dataInicio: props['Data/Hora Aceite Contrato']?.date?.start || null,
      vencimentoContrato: props['Vencimento do Contrato']?.date?.start || null,
      observacoes: props['Observações']?.rich_text?.[0]?.plain_text || '',
    };
  });
}

app.post('/portal-aluna/login/solicitar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  try {
    const matriculas = await buscarMatriculasPorCpf(cpfLimpo);
    if (matriculas.length === 0) return res.json({ ok: false, erro: 'CPF não encontrado ou sem matrícula ativa. Fale com o Fábio.' });
    const numBr = matriculas[0].contato.replace(/\D/g, '');
    await enviarOtp('aluna_' + cpfLimpo, numBr.length === 11 ? '55' + numBr : numBr, matriculas[0].nome);
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-aluna/login/solicitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-aluna/login/verificar', async (req, res) => {
  const cpfLimpo = (req.body.cpf || '').replace(/\D/g, '');
  const verificacao = verificarOtp('aluna_' + cpfLimpo, req.body.codigo);
  if (!verificacao.ok) return res.json(verificacao);
  try {
    const matriculas = await buscarMatriculasPorCpf(cpfLimpo);
    if (matriculas.length === 0) return res.json({ ok: false, erro: 'Cadastro não encontrado.' });
    const dadosLogin = { nome: matriculas[0].nome, contato: matriculas[0].contato, matriculas, cpf: cpfLimpo };
    const sessionToken = criarSessao('aluna_' + cpfLimpo, dadosLogin);
    res.json({ ok: true, ...dadosLogin, sessionToken });
  } catch (err) {
    console.error('[portal-aluna/login/verificar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-aluna/sessao/verificar', (req, res) => {
  const dados = verificarSessao(req.body.sessionToken);
  if (!dados) return res.json({ ok: false, erro: 'Sessão expirada.' });
  res.json({ ok: true, ...dados });
});

app.get('/portal-aluna/presencas/:cpf', async (req, res) => {
  const cpfLimpo = req.params.cpf.replace(/\D/g, '');
  const dadosSessaoAluna = verificarSessao(req.query.sessionToken);
  if (!dadosSessaoAluna) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAluna.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const rAlunas = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'CPF', rich_text: { equals: cpfLimpo } }, page_size: 20 }),
    });
    const dAlunas = await rAlunas.json();
    const alunaIds = (dAlunas.results || []).map(p => p.id);
    if (alunaIds.length === 0) return res.json({ ok: true, presencas: [] });

    const rPresencas = await fetch('https://api.notion.com/v1/databases/' + PRESENCAS_2026_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { or: alunaIds.map(id => ({ property: 'Aluna', relation: { contains: id } })) },
        sorts: [{ property: 'Data', direction: 'descending' }],
        page_size: 100,
      }),
    });
    const dPresencas = await rPresencas.json();
    const presencas = (dPresencas.results || []).map(p => ({
      data: p.properties['Data']?.date?.start || '',
      turma: p.properties['Turma']?.select?.name || '',
      status: p.properties['Status']?.select?.name || '',
      tipo: p.properties['Tipo']?.select?.name || '',
      professor: p.properties['Professor']?.select?.name || '',
    }));

    const totalPresencas = presencas.filter(p => p.status === 'Presente').length;
    const totalFaltas = presencas.filter(p => p.status === 'Falta').length;

    res.json({ ok: true, presencas, totalPresencas, totalFaltas });
  } catch (err) {
    console.error('[portal-aluna/presencas] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-aluna/atualizar-contato', async (req, res) => {
  const { cpf, contatoEmergenciaNome, contatoEmergenciaTelefone, possuiAlergias, quaisAlergias, usaMedicamentos, quaisMedicamentos, condicaoSaude, qualCondicao, sessionToken } = req.body;
  const cpfLimpo = (cpf || '').replace(/\D/g, '');
  const dadosSessaoAluna = verificarSessao(sessionToken);
  if (!dadosSessaoAluna) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAluna.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  if (!cpfLimpo || !contatoEmergenciaNome || !contatoEmergenciaTelefone) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'CPF', rich_text: { equals: cpfLimpo } }, page_size: 20 }),
    });
    const d = await r.json();
    const paginas = d.results || [];
    if (paginas.length === 0) return res.status(404).json({ ok: false, erro: 'Cadastro não encontrado.' });

    for (const pagina of paginas) {
      await fetch('https://api.notion.com/v1/pages/' + pagina.id, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          properties: {
            'Contato de Emergência': { rich_text: [{ text: { content: contatoEmergenciaNome } }] },
            'Tel. Emergência': { phone_number: contatoEmergenciaTelefone },
            'Possui alergias?': { select: { name: possuiAlergias || 'Não' } },
            'Quais alergias?': { rich_text: [{ text: { content: quaisAlergias || '' } }] },
            'Usa medicamentos?': { select: { name: usaMedicamentos || 'Não' } },
            'Quais medicamentos?': { rich_text: [{ text: { content: quaisMedicamentos || '' } }] },
            'Condição de saúde?': { select: { name: condicaoSaude || 'Não' } },
            'Qual condição?': { rich_text: [{ text: { content: qualCondicao || '' } }] },
          },
        }),
      });
    }
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-aluna/atualizar-contato] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

async function notificarSolicitacaoAluna(tipo, nome, contato, detalhes) {
  const msg = '📋 *Solicitação de aluna* — ' + tipo + '\n\nAluna: ' + nome + ' (' + contato + ')\n' + detalhes + '\n\n_Pendente de confirmação sua._';
  try { await enviarWhatsApp(WHATSAPP_FABIO, msg); } catch(e) {}
}

// GET /reposicao/verificar-cota/:cpf/:modalidade
app.get('/reposicao/verificar-cota/:cpf/:modalidade', async (req, res) => {
  const cpfLimpo = (req.params.cpf || '').replace(/\D/g, '');
  const modalidade = req.params.modalidade;
  if (!cpfLimpo || !modalidade) return res.status(400).json({ ok: false, erro: 'CPF e modalidade são obrigatórios.' });
  try {
    const resultado = await verificarCotaReposicao(cpfLimpo, modalidade);
    res.json({ ok: true, ...resultado });
  } catch (err) {
    console.error('[reposicao/verificar-cota] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// POST /reposicao/confirmar — gate + consumo de credito sem escrever em Alunas.
// Usado pelo Aéreos, cujo fluxo de reposição (vagas locais + WhatsApp) não passa por /inscricao.
app.post('/reposicao/confirmar', async (req, res) => {
  const { cpf, modalidade, turmaDesejada, dataDesejada } = req.body;
  const cpfLimpo = (cpf || '').replace(/\D/g, '');
  if (!cpfLimpo || !modalidade) return res.status(400).json({ ok: false, erro: 'CPF e modalidade são obrigatórios.' });
  try {
    const cotaInfo = await verificarCotaReposicao(cpfLimpo, modalidade);
    if (!cotaInfo.podeAgendar) {
      return res.status(400).json({ ok: false, erro: cotaInfo.mensagemBloqueio });
    }
    await consumirCreditoReposicao(cotaInfo.creditos[0].id, { turmaDesejada, dataDesejada });
    res.json({ ok: true });
  } catch (err) {
    console.error('[reposicao/confirmar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// GET /portal-aluna/reposicoes/:cpf — creditos abertos agrupados por modalidade, para a aba "Minha Presença"
const MAPA_MODALIDADES_REPOSICAO = {
  'Aéreos': 'https://agende-aereos.ciadoliquidificador.com.br?reposicao=1',
  'Circo Infantil': 'https://agende-infantil.ciadoliquidificador.com.br?reposicao=1',
  'Yoga': 'https://agende-yoga.ciadoliquidificador.com.br?reposicao=1',
};

app.get('/portal-aluna/reposicoes/:cpf', async (req, res) => {
  const cpfLimpo = (req.params.cpf || '').replace(/\D/g, '');
  if (!cpfLimpo) return res.status(400).json({ ok: false, erro: 'CPF é obrigatório.' });
  const dadosSessaoAluna = verificarSessao(req.query.sessionToken);
  if (!dadosSessaoAluna) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAluna.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const hoje = new Date().toISOString().slice(0, 10);
    const modalidades = [];
    for (const [modalidade, linkRemarcar] of Object.entries(MAPA_MODALIDADES_REPOSICAO)) {
      const resultado = await verificarCotaReposicao(cpfLimpo, modalidade);
      if (resultado.creditosAbertos === 0) continue;
      const creditos = resultado.creditos.map(c => ({
        dataFalta: c.dataFalta,
        prazoLimite: c.prazoLimite,
        diasRestantes: Math.ceil((new Date(c.prazoLimite + 'T12:00:00Z') - new Date(hoje + 'T12:00:00Z')) / 86400000),
      }));
      modalidades.push({ modalidade, cota: resultado.cota, creditosAbertos: resultado.creditosAbertos, creditos, linkRemarcar });
    }
    res.json({ ok: true, modalidades });
  } catch (err) {
    console.error('[portal-aluna/reposicoes] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// Descontinuado: escrevia com o schema antigo de Reposições Solicitadas (Mês de Referência,
// Status "Solicitada"/"Negada"), incompatível com o sistema de cota por janela rolante de 30 dias.
// Investigação confirmou zero chamadas em qualquer frontend (Aluna, Aéreos, Infantil, Yoga, Acro, Prof).
// Fluxo atual: apps Aéreos/Infantil/Yoga com ?reposicao=1 + POST /reposicao/confirmar.
app.post('/portal-aluna/solicitar-reposicao', (req, res) => {
  res.status(410).json({ erro: 'Este endpoint foi descontinuado. Use o fluxo de reposição dos apps (Aéreos/Infantil/Yoga) com ?reposicao=1.' });
});

app.post('/portal-aluna/solicitar-mudanca-turma', async (req, res) => {
  const { cpf, nome, contato, modalidade, turmaAtual, novaTurma, motivo, sessionToken } = req.body;
  if (!cpf || !modalidade || !novaTurma || !motivo) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios, incluindo o motivo.' });
  }
  const dadosSessaoAlunaMT = verificarSessao(sessionToken);
  if (!dadosSessaoAlunaMT) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaMT.cpf !== (cpf || '').replace(/\D/g, '')) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const dadosModalidade = MODALIDADES_MATRICULA[modalidade];
    const infoTurmaNova = dadosModalidade?.turmas.find(t => t.nome === novaTurma);
    const infoTurmaAntiga = dadosModalidade?.turmas.find(t => t.nome === turmaAtual);
    if (!infoTurmaNova) return res.status(400).json({ ok: false, erro: 'Turma inválida.' });

    const rCheck = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: { and: [
          { property: 'Modalidade', select: { equals: modalidade } },
          { property: 'Turma', select: { equals: novaTurma } },
          { property: 'Status', select: { equals: 'Ativa' } },
        ]},
        page_size: 200,
      }),
    });
    const dCheck = await rCheck.json();
    const ocupadas = (dCheck.results || []).length;
    if (ocupadas >= infoTurmaNova.limite) {
      return res.json({ ok: false, erro: 'Essa turma já está lotada. Escolha outro horário.' });
    }

    // Cria registro estruturado de aprovação pendente
    try {
      const cpfLimpoAprov = (cpf || '').replace(/\D/g, '');
      const matriculasAprov = await buscarMatriculasPorCpf(cpfLimpoAprov);
      const alvoAprov = localizarMatricula(matriculasAprov, modalidade, turmaAtual)[0] || null;
      const rAprov = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: APROVACOES_ALUNAS_DB },
          properties: {
            'Título': { title: [{ text: { content: nome + ' — Mudança de Turma' } }] },
            'Tipo': { select: { name: 'Mudança de Turma' } },
            'CPF Aluna': { rich_text: [{ text: { content: cpfLimpoAprov } }] },
            'Nome Aluna': { rich_text: [{ text: { content: nome || '' } }] },
            'Contato': { phone_number: contato || null },
            'Modalidade': { rich_text: [{ text: { content: modalidade || '' } }] },
            'Turma Atual': { rich_text: [{ text: { content: turmaAtual || '' } }] },
            'Turma Nova': { rich_text: [{ text: { content: novaTurma || '' } }] },
            'Motivo': { rich_text: [{ text: { content: motivo || '' } }] },
            'Status': { select: { name: 'Pendente' } },
            'Matrícula': alvoAprov?.pageId ? { relation: [{ id: alvoAprov.pageId }] } : undefined,
          },
        }),
      });
      if (!rAprov.ok) { const eBody = await rAprov.text(); console.error('[solicitar-mudanca-turma] Notion recusou criar aprovacao:', eBody); }
    } catch (eAprov) { console.error('[solicitar-mudanca-turma] erro ao criar aprovacao:', eAprov.message); }

    await notificarSolicitacaoAluna(
      'Mudança de turma',
      nome, contato,
      'De: ' + (turmaAtual || '(não informado)') + '\nPara: ' + novaTurma + '\nMotivo: ' + motivo
    );

    // Avisa o professor antigo e o novo (se forem diferentes)
    const msgProfessores = 'ℹ️ *Aviso de mudança de turma*\n\nAluna: ' + nome + '\nDe: ' + (turmaAtual || '(não informado)') + '\nPara: ' + novaTurma + '\nMotivo: ' + motivo;
    if (infoTurmaAntiga?.professor) {
      const telAntigo = await buscarTelefoneProfessorPorNome(infoTurmaAntiga.professor);
      if (telAntigo) { try { await enviarWhatsAppComHorarioComercial(telAntigo, msgProfessores); } catch(e) {} }
    }
    if (infoTurmaNova?.professor && infoTurmaNova.professor !== infoTurmaAntiga?.professor) {
      const telNovo = await buscarTelefoneProfessorPorNome(infoTurmaNova.professor);
      if (telNovo) { try { await enviarWhatsAppComHorarioComercial(telNovo, msgProfessores); } catch(e) {} }
    }

    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-aluna/solicitar-mudanca-turma] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-aluna/solicitar-mudanca-plano', async (req, res) => {
  const { cpf, nome, contato, modalidade, planoAtual, frequenciaAtual, novoPlano, novaFrequencia, sessionToken } = req.body;
  if (!cpf || !modalidade || !novoPlano || !novaFrequencia) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  const dadosSessaoAlunaMP = verificarSessao(sessionToken);
  if (!dadosSessaoAlunaMP) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaMP.cpf !== (cpf || '').replace(/\D/g, '')) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const dadosModalidade = MODALIDADES_MATRICULA[modalidade];
    const novoValor = dadosModalidade?.precos?.[novaFrequencia]?.[novoPlano];
    if (novoValor === undefined) return res.status(400).json({ ok: false, erro: 'Combinação de plano/frequência inválida.' });

    // Cria registro estruturado de aprovação pendente
    try {
      const cpfLimpoAprov = (cpf || '').replace(/\D/g, '');
      const matriculasAprov = await buscarMatriculasPorCpf(cpfLimpoAprov);
      const alvoAprov = localizarMatricula(matriculasAprov, modalidade)[0] || null;
      const rAprov = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: APROVACOES_ALUNAS_DB },
          properties: {
            'Título': { title: [{ text: { content: nome + ' — Mudança de Plano' } }] },
            'Tipo': { select: { name: 'Mudança de Plano' } },
            'CPF Aluna': { rich_text: [{ text: { content: cpfLimpoAprov } }] },
            'Nome Aluna': { rich_text: [{ text: { content: nome || '' } }] },
            'Contato': { phone_number: contato || null },
            'Modalidade': { rich_text: [{ text: { content: modalidade || '' } }] },
            'Plano Atual': { rich_text: [{ text: { content: planoAtual || '' } }] },
            'Plano Novo': { rich_text: [{ text: { content: novoPlano || '' } }] },
            'Frequência Atual': { rich_text: [{ text: { content: frequenciaAtual || '' } }] },
            'Frequência Nova': { rich_text: [{ text: { content: novaFrequencia || '' } }] },
            'Valor Novo': { number: novoValor },
            'Status': { select: { name: 'Pendente' } },
            'Matrícula': alvoAprov?.pageId ? { relation: [{ id: alvoAprov.pageId }] } : undefined,
          },
        }),
      });
      if (!rAprov.ok) { const eBody = await rAprov.text(); console.error('[solicitar-mudanca-plano] Notion recusou criar aprovacao:', eBody); }
    } catch (eAprov) { console.error('[solicitar-mudanca-plano] erro ao criar aprovacao:', eAprov.message); }

    await notificarSolicitacaoAluna(
      'Mudança de plano',
      nome, contato,
      'De: ' + (planoAtual || '?') + ' (' + (frequenciaAtual || '?') + ')\nPara: ' + novoPlano + ' (' + novaFrequencia + ') — novo valor estimado: R$ ' + novoValor.toFixed(2)
    );
    res.json({ ok: true, novoValor });
  } catch (err) {
    console.error('[portal-aluna/solicitar-mudanca-plano] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-aluna/atualizar-cadastro', async (req, res) => {
  const { cpf, nome, email, contato, endereco, sessionToken } = req.body;
  const cpfLimpo = (cpf || '').replace(/\D/g, '');
  const dadosSessaoAlunaAC = verificarSessao(sessionToken);
  if (!dadosSessaoAlunaAC) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaAC.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  if (!cpfLimpo || !nome || !email || !contato) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'CPF', rich_text: { equals: cpfLimpo } }, page_size: 20 }),
    });
    const d = await r.json();
    const paginas = d.results || [];
    if (paginas.length === 0) return res.status(404).json({ ok: false, erro: 'Cadastro não encontrado.' });
    for (const pagina of paginas) {
      await fetch('https://api.notion.com/v1/pages/' + pagina.id, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          properties: {
            'Nome': { title: [{ text: { content: nome } }] },
            'Email': { email: email },
            'Contato': { phone_number: contato },
            'Endereço': { rich_text: [{ text: { content: endereco || '' } }] },
          },
        }),
      });
    }
    try { await enviarWhatsApp(WHATSAPP_FABIO, 'ℹ️ *Dados cadastrais atualizados* — ' + nome + ' (' + contato + ')'); } catch(e) {}
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-aluna/atualizar-cadastro] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});


const MURAL_DB = 'c45786e213ff463f8558054b2f787a69';

app.post('/mural/login/solicitar', async (req, res) => {
  try {
    await enviarOtp('mural_fabio', WHATSAPP_FABIO, 'Fábio');
    res.json({ ok: true });
  } catch (err) {
    console.error('[mural/login/solicitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/mural/login/verificar', (req, res) => {
  const verificacao = verificarOtp('mural_fabio', req.body.codigo);
  res.json(verificacao);
});

// ===== PORTAL ADMIN — LOGIN OTP + SESSÃO =====
app.post('/portal-admin/login/solicitar', async (req, res) => {
  try {
    await enviarOtp('portal_admin_fabio', WHATSAPP_FABIO, 'Fábio');
    res.json({ ok: true });
  } catch (err) {
    console.error('[portal-admin/login/solicitar] erro:', err.message);
    res.status(500).json({ ok: false, erro: 'Erro ao enviar código.' });
  }
});

app.post('/portal-admin/login/verificar', (req, res) => {
  const verificacao = verificarOtp('portal_admin_fabio', req.body.codigo);
  if (!verificacao.ok) {
    return res.status(401).json(verificacao);
  }
  const token = criarSessao('portal_admin_fabio', {});
  res.json({ ok: true, token });
});

app.post('/portal-admin/sessao/verificar', (req, res) => {
  const token = req.body.token || '';
  const dados = verificarSessao(token);
  if (!dados) {
    return res.status(401).json({ ok: false, erro: 'Sessão inválida ou expirada.' });
  }
  res.json({ ok: true });
});
// ===== FIM PORTAL ADMIN — LOGIN OTP + SESSÃO =====



// Turmas que não usam o fluxo genérico de matrícula (MODALIDADES_MATRICULA),
// mas ainda assim gravam no banco Alunas (dual-write) e podem receber avisos do Mural.
const TURMAS_MURAL_EXTRAS = [
  { modalidade: 'Percussão Coletiva', turma: 'Terça 19h30', dia: 'Terça', horario: '19:30' },
  { modalidade: 'Dancas Brasileiras', turma: 'Quarta 20h', dia: 'Quarta', horario: '20:00' },
  { modalidade: 'Commedia Dell Arte', turma: 'Sexta 10h', dia: 'Sexta', horario: '10:00' },
];

app.get('/mural/turmas-disponiveis', (req, res) => {
  const turmas = [];
  for (const [modalidade, dados] of Object.entries(MODALIDADES_MATRICULA)) {
    dados.turmas.forEach(t => turmas.push({ modalidade, turma: t.nome, dia: t.dia, horario: t.horario }));
  }
  turmas.push(...TURMAS_MURAL_EXTRAS);
  res.json({ ok: true, turmas });
});

app.post('/mural/postar', async (req, res) => {
  const { titulo, mensagem, turmasAlvo, imagemBase64, imagemNomeArquivo } = req.body;
  if (!titulo || !mensagem || !Array.isArray(turmasAlvo) || turmasAlvo.length === 0) {
    return res.status(400).json({ ok: false, erro: 'Preencha título, mensagem e ao menos uma turma.' });
  }
  try {
    const filtrosOr = turmasAlvo.map(t => ({ and: [
      { property: 'Modalidade', select: { equals: t.modalidade } },
      { property: 'Turma', select: { equals: t.turma } },
      { property: 'Status', select: { equals: 'Ativa' } },
    ]}));
    const rAlunas = await fetch('https://api.notion.com/v1/databases/' + ALUNAS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { or: filtrosOr }, page_size: 200 }),
    });
    const dAlunas = await rAlunas.json();
    const contatos = new Set();
    (dAlunas.results || []).forEach(p => {
      const contato = p.properties['Contato']?.phone_number;
      if (contato) contatos.add(contato.replace(/\D/g, ''));
    });

    // Se veio imagem, sobe pro OneDrive e pega um link de download direto
    // (necessário pro Digisac conseguir buscar e anexar o arquivo de verdade).
    let fileUrl = null;
    if (imagemBase64) {
      try {
        const msToken = await getMicrosoftToken();
        const nomePasta = 'Mural-Imagens';
        const folderId = await criarOuObterSubpasta(msToken, nomePasta);
        const nomeArquivo = (Date.now() + '_' + (imagemNomeArquivo || 'imagem.jpg')).replace(/[^a-zA-Z0-9._-]/g, '-');
        const uploaded = await uploadFotoParaPasta(msToken, folderId, imagemBase64, nomeArquivo);
        fileUrl = uploaded['@microsoft.graph.downloadUrl'] || null;
      } catch (eImg) {
        console.error('[mural/postar] erro ao subir imagem:', eImg.message);
      }
    }

    const msgCompleta = '📢 *' + titulo + '*\n\n' + mensagem;
    let enviados = 0;
    let falhas = 0;
    for (const numLimpo of contatos) {
      const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
      try {
        if (fileUrl) {
          await enviarWhatsAppComImagem(numBr, msgCompleta, fileUrl);
        } else {
          await enviarWhatsAppComHorarioComercial(numBr, msgCompleta);
        }
        enviados++;
      } catch(e) {
        falhas++;
        console.error('[mural/postar] falha ao enviar para ' + numBr + ':', e.message);
      }
    }
    if (falhas > 0) {
      console.error('[mural/postar] resumo: ' + enviados + ' enviados, ' + falhas + ' falhas, ' + contatos.size + ' contatos encontrados. fileUrl=' + (fileUrl || '(sem imagem)'));
    }

    await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: MURAL_DB },
        properties: {
          'Título': { title: [{ text: { content: titulo } }] },
          'Mensagem': { rich_text: [{ text: { content: mensagem } }] },
          'Turmas Alvo JSON': { rich_text: [{ text: { content: JSON.stringify(turmasAlvo) } }] },
          'Autor': { rich_text: [{ text: { content: 'Fábio' } }] },
          'Enviado': { checkbox: true },
          'Total Destinatarios': { number: enviados },
        },
      }),
    });

    res.json({ ok: true, totalDestinatarios: enviados, contatosEncontrados: contatos.size, falhas });
  } catch (err) {
    console.error('[mural/postar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.get('/mural/listar/:modalidade/:turma', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + MURAL_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ sorts: [{ timestamp: 'created_time', direction: 'descending' }], page_size: 30 }),
    });
    const d = await r.json();
    const modalidade = decodeURIComponent(req.params.modalidade);
    const turma = decodeURIComponent(req.params.turma);
    const avisos = (d.results || []).filter(p => {
      let turmasAlvo = [];
      try { turmasAlvo = JSON.parse(p.properties['Turmas Alvo JSON']?.rich_text?.[0]?.plain_text || '[]'); } catch(e) {}
      return turmasAlvo.some(t => t.modalidade === modalidade && t.turma === turma);
    }).map(p => ({
      titulo: p.properties['Título']?.title?.[0]?.plain_text || '',
      mensagem: p.properties['Mensagem']?.rich_text?.[0]?.plain_text || '',
      data: p.created_time,
    }));
    res.json({ ok: true, avisos });
  } catch (err) {
    console.error('[mural/listar] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// Encontra, dentre as matriculas ativas de um CPF, a(s) que casam com modalidade/turma informadas
// (turma e opcional — se omitida, casa qualquer turma daquela modalidade).
function localizarMatricula(matriculas, modalidade, turma) {
  return matriculas.filter(m => m.modalidade === modalidade && (!turma || m.turma === turma));
}

function formatarDetalhesMulta(calculo) {
  if (!calculo || !calculo.ok) return 'Não foi possível calcular a multa automaticamente — confira manualmente.';
  if (calculo.multa > 0) {
    return 'Multa de cancelamento: R$ ' + calculo.multa.toFixed(2) +
      ' (meses cursados: ' + calculo.mesesCursados + '/' + calculo.duracaoFidelidadeMeses +
      '; valor mensal mesma frequência: R$ ' + calculo.valorMensalMesmaFrequencia.toFixed(2) +
      '; valor do plano contratado: R$ ' + calculo.valorPlanoContratado.toFixed(2) + ')';
  }
  return 'Sem multa de cancelamento' + (calculo.motivo ? (' — ' + calculo.motivo) : '') + '.';
}

// GET /portal-aluna/simular-cancelamento/:cpf?modalidade=...&turma=...
// Retorna a multa (ou "sem multa") calculada por calcularMultaCancelamento, para o Portal Aluna
// exibir para a aluna ANTES de ela confirmar o cancelamento. Sem filtro de modalidade/turma, retorna
// a simulação de todas as matrículas ativas daquele CPF.
app.get('/portal-aluna/simular-cancelamento/:cpf', async (req, res) => {
  const cpfLimpo = req.params.cpf.replace(/\D/g, '');
  const { modalidade, turma } = req.query;
  const dadosSessaoAlunaSC = verificarSessao(req.query.sessionToken);
  if (!dadosSessaoAlunaSC) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaSC.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const matriculas = await buscarMatriculasPorCpf(cpfLimpo);
    if (matriculas.length === 0) return res.json({ ok: false, erro: 'Cadastro não encontrado ou sem matrícula ativa.' });

    const alvos = modalidade ? localizarMatricula(matriculas, modalidade, turma) : matriculas;
    if (alvos.length === 0) return res.json({ ok: false, erro: 'Matrícula não encontrada para os filtros informados.' });

    const simulacoes = alvos.map(m => {
      const calculo = calcularMultaCancelamento(m.modalidade, m.frequencia, m.plano, m.dataInicio);
      return {
        modalidade: m.modalidade, turma: m.turma, plano: m.plano, frequencia: m.frequencia,
        valorAtual: m.valor, dataInicio: m.dataInicio,
        ...calculo,
        detalhes: formatarDetalhesMulta(calculo),
      };
    });

    res.json({ ok: true, simulacoes });
  } catch (err) {
    console.error('[portal-aluna/simular-cancelamento] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// GET /portal-aluna/simular-ferias/:cpf?modalidade=...&turma=...
// Retorna, por matricula, quantos dias de ferias a aluna ainda tem disponiveis no ciclo atual.
app.get('/portal-aluna/simular-ferias/:cpf', async (req, res) => {
  const cpfLimpo = req.params.cpf.replace(/\D/g, '');
  const { modalidade, turma } = req.query;
  const dadosSessaoAlunaSF = verificarSessao(req.query.sessionToken);
  if (!dadosSessaoAlunaSF) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaSF.cpf !== cpfLimpo) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  try {
    const matriculas = await buscarMatriculasPorCpf(cpfLimpo);
    if (matriculas.length === 0) return res.json({ ok: false, erro: 'Cadastro não encontrado ou sem matrícula ativa.' });

    const alvos = modalidade ? localizarMatricula(matriculas, modalidade, turma) : matriculas;
    if (alvos.length === 0) return res.json({ ok: false, erro: 'Matrícula não encontrada para os filtros informados.' });

    const simulacoes = alvos.map(m => {
      const diasPorCiclo = FERIAS_DIAS_POR_PLANO[m.plano] || 0;
      if (diasPorCiclo === 0 || !m.dataInicio) {
        return {
          modalidade: m.modalidade, turma: m.turma, plano: m.plano,
          diasPorCiclo: 0, diasUsados: 0, diasDisponiveis: 0,
          motivo: !m.dataInicio ? 'Data de início da matrícula não encontrada.' : 'Plano ' + m.plano + ' não possui direito a férias.',
        };
      }
      const periodos = extrairPeriodosFerias(m.observacoes);
      const ciclo = calcularCicloAtualFerias(m.dataInicio, m.plano, periodos);
      const diasUsados = diasFeriasUsadosNoCiclo(periodos, ciclo.inicioCiclo, ciclo.fimCiclo);
      const diasDisponiveis = Math.max(0, diasPorCiclo - diasUsados);
      return {
        modalidade: m.modalidade, turma: m.turma, plano: m.plano,
        diasPorCiclo, diasUsados, diasDisponiveis,
        inicioCiclo: ciclo.inicioCiclo.toISOString(), fimCiclo: ciclo.fimCiclo.toISOString(),
        avisoMinimoDias: FERIAS_AVISO_MINIMO_DIAS,
        motivo: null,
      };
    });

    res.json({ ok: true, simulacoes });
  } catch (err) {
    console.error('[portal-aluna/simular-ferias] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// POST /portal-aluna/solicitar-ferias
// Registra um periodo de ferias, prorroga o Vencimento do Contrato automaticamente e
// notifica o Fabio. Nao depende de aprovacao formal (e um direito ja previsto em contrato).
app.post('/portal-aluna/solicitar-ferias', async (req, res) => {
  const { cpf, nome, contato, modalidade, turma, dataInicioFerias, dataFimFerias, sessionToken } = req.body;
  if (!cpf || !modalidade || !dataInicioFerias || !dataFimFerias) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  const cpfLimpoSF = (cpf || '').replace(/\D/g, '');
  const dadosSessaoAlunaSF2 = verificarSessao(sessionToken);
  if (!dadosSessaoAlunaSF2) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaSF2.cpf !== cpfLimpoSF) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });

  try {
    const inicioFerias = new Date(dataInicioFerias);
    const fimFerias = new Date(dataFimFerias);
    if (isNaN(inicioFerias.getTime()) || isNaN(fimFerias.getTime())) {
      return res.status(400).json({ ok: false, erro: 'Datas inválidas.' });
    }
    if (fimFerias < inicioFerias) {
      return res.status(400).json({ ok: false, erro: 'A data final não pode ser anterior à data inicial.' });
    }

    const hojeMaisAviso = apenasData(new Date());
    hojeMaisAviso.setDate(hojeMaisAviso.getDate() + FERIAS_AVISO_MINIMO_DIAS);
    if (apenasData(inicioFerias) < hojeMaisAviso) {
      return res.status(400).json({
        ok: false,
        erro: 'As férias precisam ser solicitadas com pelo menos ' + FERIAS_AVISO_MINIMO_DIAS + ' dias de antecedência.',
      });
    }

    const matriculas = await buscarMatriculasPorCpf(cpfLimpoSF);
    const alvo = localizarMatricula(matriculas, modalidade, turma)[0] || null;
    if (!alvo) return res.status(404).json({ ok: false, erro: 'Matrícula não encontrada.' });
    if (!alvo.dataInicio) return res.status(400).json({ ok: false, erro: 'Data de início da matrícula não encontrada — fale com a gente.' });

    const diasPorCiclo = FERIAS_DIAS_POR_PLANO[alvo.plano] || 0;
    if (diasPorCiclo === 0) {
      return res.status(400).json({ ok: false, erro: 'Plano ' + alvo.plano + ' não possui direito a férias.' });
    }

    const periodosExistentes = extrairPeriodosFerias(alvo.observacoes);
    if (periodoFeriasSobrepoe(periodosExistentes, dataInicioFerias, dataFimFerias)) {
      return res.status(400).json({ ok: false, erro: 'Esse período se sobrepõe a férias já registradas.' });
    }

    const diasSolicitados = diasNoIntervaloFerias(dataInicioFerias, dataFimFerias);
    const ciclo = calcularCicloAtualFerias(alvo.dataInicio, alvo.plano, periodosExistentes, inicioFerias);
    const diasUsados = diasFeriasUsadosNoCiclo(periodosExistentes, ciclo.inicioCiclo, ciclo.fimCiclo);
    const diasDisponiveis = diasPorCiclo - diasUsados;

    if (diasSolicitados > diasDisponiveis) {
      return res.status(400).json({
        ok: false,
        erro: 'Saldo insuficiente: restam ' + diasDisponiveis + ' dia(s) de férias neste ciclo (solicitado: ' + diasSolicitados + ').',
        diasDisponiveis,
      });
    }

    const registroTexto = 'FERIAS: ' + dataInicioFerias + ' a ' + dataFimFerias +
      ' (' + diasSolicitados + ' dias, solicitado em ' + new Date().toLocaleString('pt-BR') + ')';
    const observacoesAtualizadas = (alvo.observacoes ? alvo.observacoes + '\n\n' : '') + registroTexto;

    const vencimentoBase = alvo.vencimentoContrato
      ? new Date(alvo.vencimentoContrato)
      : new Date(calcularVencimentoContrato(alvo.dataInicio, alvo.plano));
    const novoVencimento = new Date(vencimentoBase.getTime());
    novoVencimento.setDate(novoVencimento.getDate() + diasSolicitados);

    const rPatchFerias = await fetch('https://api.notion.com/v1/pages/' + alvo.pageId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        properties: {
          'Observações': { rich_text: [{ text: { content: observacoesAtualizadas.slice(-1900) } }] },
          'Vencimento do Contrato': { date: { start: novoVencimento.toISOString() } },
        },
      }),
    });
    if (!rPatchFerias.ok) {
      const eBody = await rPatchFerias.text();
      console.error('[portal-aluna/solicitar-ferias] Notion recusou atualizar matricula:', eBody);
      return res.status(500).json({ ok: false, erro: 'Notion recusou gravar as férias — ver logs.' });
    }

    try {
      await enviarWhatsApp(
        WHATSAPP_FABIO,
        '🏖️ *Férias registradas*\n\nAluna: ' + (nome || alvo.nome || '') +
        '\nModalidade: ' + modalidade + (turma ? ' — ' + turma : '') +
        '\nPeríodo: ' + dataInicioFerias + ' a ' + dataFimFerias + ' (' + diasSolicitados + ' dias)' +
        '\nNovo vencimento do contrato: ' + novoVencimento.toLocaleDateString('pt-BR')
      );
    } catch (eNotif) {
      console.error('[portal-aluna/solicitar-ferias] erro ao notificar Fábio:', eNotif.message);
    }

    res.json({
      ok: true,
      diasSolicitados,
      diasDisponiveisRestantes: diasDisponiveis - diasSolicitados,
      novoVencimento: novoVencimento.toISOString(),
    });
  } catch (err) {
    console.error('[portal-aluna/solicitar-ferias] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/portal-aluna/solicitar-cancelamento', async (req, res) => {
  const { cpf, nome, contato, modalidade, turma, motivo, sessionToken } = req.body;
  if (!cpf || !modalidade) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  const dadosSessaoAlunaSCa = verificarSessao(sessionToken);
  if (!dadosSessaoAlunaSCa) return res.status(401).json({ ok: false, erro: 'Sessão expirada. Faça login novamente.' });
  if (dadosSessaoAlunaSCa.cpf !== (cpf || '').replace(/\D/g, '')) return res.status(403).json({ ok: false, erro: 'Acesso negado.' });
  const cpfLimpo = (cpf || '').replace(/\D/g, '');
  try {
    const matriculas = await buscarMatriculasPorCpf(cpfLimpo);
    const alvos = localizarMatricula(matriculas, modalidade, turma);
    const matricula = alvos[0] || null;

    let calculo = null;
    if (matricula) {
      calculo = calcularMultaCancelamento(matricula.modalidade, matricula.frequencia, matricula.plano, matricula.dataInicio);

      // Registra a solicitacao e a trilha de auditoria do calculo da multa direto na matricula, no Notion.
      if (matricula.pageId) {
        const registroCancelamento =
          'Cancelamento solicitado em ' + new Date().toLocaleString('pt-BR') + '. Motivo: ' + (motivo || '(não informado)') +
          (calculo && calculo.ok ? ('\n' + formatarDetalhesMulta(calculo)) : '');
        const observacoesAtualizadas = (matricula.observacoes ? matricula.observacoes + '\n\n' : '') + registroCancelamento;
        try {
          await fetch('https://api.notion.com/v1/pages/' + matricula.pageId, {
            method: 'PATCH',
            headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
            body: JSON.stringify({ properties: { 'Observações': { rich_text: paraRichTextEmBlocos(observacoesAtualizadas) } } }),
          });
        } catch (e) { console.error('[portal-aluna/solicitar-cancelamento] erro ao gravar no Notion:', e.message); }
      }
    }

    // Cria registro estruturado de aprovação pendente
    try {
      const rAprov = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parent: { database_id: APROVACOES_ALUNAS_DB },
          properties: {
            'Título': { title: [{ text: { content: nome + ' — Cancelamento' } }] },
            'Tipo': { select: { name: 'Cancelamento' } },
            'CPF Aluna': { rich_text: [{ text: { content: cpfLimpo } }] },
            'Nome Aluna': { rich_text: [{ text: { content: nome || '' } }] },
            'Contato': { phone_number: contato || null },
            'Modalidade': { rich_text: [{ text: { content: modalidade || '' } }] },
            'Turma Atual': { rich_text: [{ text: { content: turma || '' } }] },
            'Motivo': { rich_text: [{ text: { content: motivo || '' } }] },
            'Detalhes Multa': { rich_text: [{ text: { content: formatarDetalhesMulta(calculo) } }] },
            'Status': { select: { name: 'Pendente' } },
            'Matrícula': matricula?.pageId ? { relation: [{ id: matricula.pageId }] } : undefined,
          },
        }),
      });
      if (!rAprov.ok) { const eBody = await rAprov.text(); console.error('[solicitar-cancelamento] Notion recusou criar aprovacao:', eBody); }
    } catch (eAprov) { console.error('[solicitar-cancelamento] erro ao criar aprovacao:', eAprov.message); }

    await notificarSolicitacaoAluna(
      'Cancelamento de contrato',
      nome, contato,
      'Modalidade: ' + modalidade + (turma ? (' — ' + turma) : '') + '\nMotivo: ' + (motivo || '(não informado)') +
      '\n\n⚠️ Contrato prevê 30 dias de aviso prévio a partir desta comunicação escrita — confirmar data de encerramento com a aluna.' +
      '\n💰 ' + formatarDetalhesMulta(calculo)
    );
    res.json({ ok: true, calculoMulta: calculo });
  } catch (err) {
    console.error('[portal-aluna/solicitar-cancelamento] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// ============================================================
// OTP GENERICO — senha unica enviada por WhatsApp (Portal Aluna, Portal Profs, Mural)
// ============================================================
const otpStore = {};
const sessaoStore = {};

function criarSessao(identificador, dados) {
  const token = require('crypto').randomBytes(16).toString('hex');
  sessaoStore[token] = { identificador, dados, expiraEm: Date.now() + 10 * 60000 };
  return token;
}

function verificarSessao(token) {
  const registro = sessaoStore[token];
  if (!registro) return null;
  if (Date.now() > registro.expiraEm) { delete sessaoStore[token]; return null; }
  return registro.dados;
}

function gerarCodigoOtp() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

async function enviarOtp(identificador, numeroWhatsapp, nome) {
  const codigo = gerarCodigoOtp();
  otpStore[identificador] = { codigo, expiraEm: Date.now() + 10 * 60000 };
  await enviarWhatsApp(numeroWhatsapp, 'Olá' + (nome ? ', ' + nome.split(' ')[0] : '') + '! 🔐\n\nSeu código de acesso é: *' + codigo + '*\n\nVálido por 10 minutos.');
}

function verificarOtp(identificador, codigoDigitado) {
  const registro = otpStore[identificador];
  if (!registro) return { ok: false, erro: 'Código não solicitado ou expirado. Peça um novo código.' };
  if (Date.now() > registro.expiraEm) { delete otpStore[identificador]; return { ok: false, erro: 'Código expirado. Peça um novo código.' }; }
  if (registro.codigo !== String(codigoDigitado || '').trim()) return { ok: false, erro: 'Código incorreto.' };
  delete otpStore[identificador];
  return { ok: true };
}

async function buscarTelefoneProfessorPorNome(nome) {
  try {
    const telNotion = await buscarProfessorPorNome(nome);
    if (telNotion) return telNotion;
  } catch (e) {}
  const prof = PROFESSORES_PORTAL.find(p => p.nome === nome);
  return prof ? prof.telefone : null;
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

async function nomeTituloDaPaginaRelatorio(pageId) {
  try {
    const rp = await fetch('https://api.notion.com/v1/pages/' + pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    if (!rp.ok) return '';
    const pd = await rp.json();
    for (const key in (pd.properties || {})) {
      if (pd.properties[key].type === 'title') return (pd.properties[key].title?.[0]?.plain_text || '').trim();
    }
    return '';
  } catch (e) { return ''; }
}

app.post('/relatorio-apresentacao', async (req, res) => {
  const { notionPageId, local, localNome, horario, data, produtor, publico, intercorrencia, fotos, valorAlimentacaoTotal, detalheAlimentacao, notasFiscais } = req.body;
  if (!notionPageId || !produtor || publico === undefined) return res.status(400).json({ error: 'Campos obrigatórios faltando' });
  try {
    const msToken = await getMicrosoftToken();
    const localSlug = slugify(localNome || local);
    const produtorSlug = slugify(produtor);
    const fotosValidas = (fotos || []).filter(Boolean);
    const notasFiscaisValidas = (notasFiscais || []).filter(Boolean);

    let trabalhoNome = '';
    try {
      const rPagina = await fetch('https://api.notion.com/v1/pages/' + notionPageId, {
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
      });
      const paginaData = await rPagina.json();
      const trabalhoRel = paginaData.properties?.['🎭 Trabalhos']?.relation || [];
      if (trabalhoRel.length) trabalhoNome = await nomeTituloDaPaginaRelatorio(trabalhoRel[0].id);
    } catch (e) { console.error('[produtor] erro ao buscar nome do trabalho:', e.message); }
    const trabalhoSlug = slugify(trabalhoNome || 'apresentacao');

    // Formata a data de YYYY-MM-DD para DD_MM_AA
    let dataFormatada = data;
    try {
      const [ano, mes, dia] = (data || '').split('-');
      if (ano && mes && dia) dataFormatada = `${dia}_${mes}_${ano.slice(2)}`;
    } catch (e) {}

    // Pasta principal: "Nome do Trabalho - DD_MM_AA"
    const nomePastaPrincipal = `${trabalhoNome || 'Apresentação'} - ${dataFormatada}`;
    const pastaPrincipalId = await criarOuObterSubpasta(msToken, nomePastaPrincipal);

    // Subpastas dentro da principal: "fotos" e "notas fiscais"
    const pastaFotosId = await criarSubpastaDentroDe(msToken, pastaPrincipalId, 'fotos');
    const pastaNotasFiscaisId = await criarSubpastaDentroDe(msToken, pastaPrincipalId, 'notas fiscais');

    const labelsFotos = ['inicio', 'meio', 'fim', 'publico'];
    const uploadPromises = fotosValidas.map((foto, i) => {
      const filename = `${data}_${trabalhoSlug}_${localSlug}_${produtorSlug}_${labelsFotos[i] || 'foto' + (i+1)}.jpg`;
      return uploadFotoParaPasta(msToken, pastaFotosId, foto, filename);
    });
    await Promise.all(uploadPromises);

    const uploadNotasFiscaisPromises = notasFiscaisValidas.map((nf, i) => {
      const filename = `${data}_${trabalhoSlug}_${localSlug}_${produtorSlug}_nota-fiscal-${i + 1}.jpg`;
      return uploadFotoParaPasta(msToken, pastaNotasFiscaisId, nf, filename);
    });
    await Promise.all(uploadNotasFiscaisPromises);

    // Link compartilhado aponta pra pasta PRINCIPAL (contém fotos/ e notas fiscais/ dentro)
    const linkPasta = await criarLinkCompartilhamento(msToken, pastaPrincipalId);
    const links = [linkPasta];
    const notionResp = await fetch(`https://api.notion.com/v1/pages/${notionPageId}`, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: {
        'Público': { number: parseInt(publico) },
        'Intercorrência': { rich_text: [{ text: { content: intercorrencia || 'Nenhuma' } }] },
        'Fotos Drive': { url: links[0] || null },
        'Valor Alimentacao': { number: valorAlimentacaoTotal !== undefined && valorAlimentacaoTotal !== null ? Number(valorAlimentacaoTotal) : null },
        'Detalhe Alimentacao': { rich_text: detalheAlimentacao ? [{ text: { content: detalheAlimentacao } }] : [] },
        'Notas Fiscais Drive': { url: notasFiscaisValidas.length ? links[0] : null },
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
  const { nome, whatsapp, email, turmaId, data, tipo, cpf } = req.body;
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

  let creditoReposicaoId = null;
  if (tipo === 'reposicao') {
    if (!cpf) return res.status(400).json({ ok: false, error: 'CPF é obrigatório para agendar reposição.' });
    const cpfLimpo = cpf.replace(/\D/g, '');
    const cotaInfo = await verificarCotaReposicao(cpfLimpo, 'Yoga');
    if (!cotaInfo.podeAgendar) {
      return res.status(400).json({ ok: false, error: cotaInfo.mensagemBloqueio });
    }
    creditoReposicaoId = cotaInfo.creditos[0].id;
  }

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

    if (creditoReposicaoId) {
      await consumirCreditoReposicao(creditoReposicaoId, { turmaDesejada: turma.nome, dataDesejada: data });
    }

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
    nome: 'André',
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
      .map(p => {
        const observacoesAluna = p.properties?.Observações?.rich_text?.[0]?.plain_text || '';
        const periodosFeriasAluna = extrairPeriodosFerias(observacoesAluna);
        const emFerias = data ? dataEstaEmFerias(periodosFeriasAluna, data) : false;
        return {
          id: p.id,
          nome: p.properties?.Nome?.title?.[0]?.plain_text || '',
          tipo: emFerias ? 'Ferias' : 'Regular',
        };
      })
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

    // Cota de reposicao: toda Falta Regular gera (ou ja tem) um credito com prazo de 30 dias
    const faltasParaCredito = criados.filter(c => c.tipo !== 'Experimental' && c.status === 'Falta' && c.pageId && c.alunaId);
    for (const falta of faltasParaCredito) {
      await criarCreditoReposicaoPorFalta({
        faltaPageId: falta.pageId,
        alunaId: falta.alunaId,
        nomeAluna: falta.nome,
        modalidade,
        turma,
        dataFalta: data,
      });
    }

    const experimentaisRegistradas = presencas.filter(p => p.tipo === 'Experimental');
    for (const exp of experimentaisRegistradas) {
      const resultado = exp.status === 'Presente' ? 'COMPARECEU ✅' : 'FALTOU ❌';

      // Espelha o resultado direto no banco Alunas, pra facilitar visualizacao
      if (exp.alunaId) {
        try {
          await fetch('https://api.notion.com/v1/pages/' + exp.alunaId, {
            method: 'PATCH',
            headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
            body: JSON.stringify({ properties: { 'Compareceu Experimental': { select: { name: exp.status === 'Presente' ? 'Sim' : 'Nao' } } } }),
          });
        } catch (eCompareceu) {
          console.error('[presenca] erro ao marcar Compareceu Experimental:', eCompareceu.message);
        }
      }
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
              CONVERSAS_ESTADO[numeroAluna] = { estado: 'aguardando_sim_nao_matricula', alunaId: exp.alunaId, nome: exp.nome, modalidade };
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

    if (estado.estado === 'aguardando_resposta_sub') {
      const broadcast = SUBSTITUICOES_BROADCAST[estado.broadcastId];
      if (!broadcast || broadcast.resolvido) {
        await enviarWhatsApp(numero, 'Essa substituição já foi resolvida por outro professor. Obrigado por responder! 💛');
        delete CONVERSAS_ESTADO[numero];
        return;
      }
      const resposta = interpretarSimNao(texto);
      if (resposta === 'sim') {
        broadcast.resolvido = true;
        try {
          await atualizarRegistroSubstituicao(broadcast.notionPageId, {
            status: 'Resolvido - Confirmado por Broadcast', whatsappSubstituto: numero,
          });
        } catch(e) {}
        await enviarWhatsApp(numero, 'Show, muito obrigado(a)! ✅\n\nVocê está confirmado(a) na turma de ' + broadcast.turma + ' no dia ' + broadcast.dataFmt + '.\n\n💰 Valor: R$ 80,00 por aula, pago pelo(a) ' + broadcast.professorFaltante + ' direto pra você.');
        const msgFabio = '✅ *Substituição resolvida (broadcast)*\n\nProfessor: ' + broadcast.professorFaltante + '\nTurma: ' + broadcast.turma + ' (' + broadcast.modalidade + ')\nData: ' + broadcast.dataFmt + '\nSubstituto: ' + numero;
        try { await enviarWhatsApp(WHATSAPP_FABIO, msgFabio); } catch(e) {}

        // Avisa o professor titular quem vai cobrir, e lembra do pagamento
        try {
          const telTitular = await buscarTelefoneProfessorPorNome(broadcast.professorFaltante);
          if (telTitular) {
            const msgTitular = 'Boa notícia! ✅\n\nSua substituição na turma de ' + broadcast.turma + ' (' + broadcast.modalidade + ') no dia ' + broadcast.dataFmt + ' foi resolvida.\n\n💰 Lembrete: o valor de R$ 80,00 por aula deve ser pago por você diretamente ao(à) professor(a) substituto(a).';
            await enviarWhatsApp(telTitular, msgTitular);
          }
        } catch (e) {}
        const nomeSubstitutoConfirmado = (PROFESSORES_SUB[broadcast.modalidade] || []).find(p => p.telefone === numero)?.nome || 'Professor(a)';
        agendarLembreteSub(numero, nomeSubstitutoConfirmado, broadcast.professorFaltante, broadcast.turma, broadcast.modalidade, broadcast.data).catch(()=>{});
        broadcast.telefonesConsultados.forEach(tel => {
          if (tel !== numero && CONVERSAS_ESTADO[tel]?.estado === 'aguardando_resposta_sub' && CONVERSAS_ESTADO[tel]?.broadcastId === estado.broadcastId) {
            enviarWhatsApp(tel, 'Essa turma já foi coberta por outro professor. Obrigado por topar! 💛').catch(()=>{});
            delete CONVERSAS_ESTADO[tel];
          }
        });
        delete CONVERSAS_ESTADO[numero];
      } else if (resposta === 'nao') {
        if (!broadcast.recusas.includes(numero)) broadcast.recusas.push(numero);
        delete CONVERSAS_ESTADO[numero];
        if (broadcast.recusas.length >= broadcast.telefonesConsultados.length) {
          try {
            await atualizarRegistroSubstituicao(broadcast.notionPageId, { status: 'Sem Substituto - Fábio Notificado' });
          } catch(e) {}
          const msgFabio = '⚠️ *Ninguém pôde cobrir a substituição*\n\nProfessor: ' + broadcast.professorFaltante + '\nTurma: ' + broadcast.turma + ' (' + broadcast.modalidade + ')\nData: ' + broadcast.dataFmt;
          try { await enviarWhatsApp(WHATSAPP_FABIO, msgFabio); } catch(e) {}
        }
      } else {
        await enviarWhatsApp(numero, 'Desculpa, não entendi 🙏\n\nVocê pode cobrir essa aula? Responde *SIM* ou *NÃO*.');
      }
      return;
    }

    if (estado.estado === 'aguardando_sim_nao_matricula') {
      const respostaMatricula = interpretarSimNao(texto);
      if (respostaMatricula === 'sim') {
        await enviarWhatsApp(numero, 'Que ótimo! 🎉\n\nÉ só clicar no link abaixo pra fazer sua matrícula:\nhttps://matricula.ciadoliquidificador.com.br');
        delete CONVERSAS_ESTADO[numero];
      } else if (respostaMatricula === 'nao') {
        await enviarWhatsApp(numero, msgMotivo());
        CONVERSAS_ESTADO[numero] = { ...estado, estado: 'aguardando_motivo' };
      } else {
        await enviarWhatsApp(numero, msgNaoEntendiSimNao());
      }
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

// Calcula a data da Pascoa (algoritmo de Meeus/Jones/Butcher), usada para
// derivar os feriados moveis (Carnaval, Sexta-feira Santa, Corpus Christi).
function calcularPascoa(ano) {
  const a = ano % 19;
  const b = Math.floor(ano / 100);
  const c = ano % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(ano, mes - 1, dia));
}

function formatarDataStrUTC(data) {
  return data.toISOString().split('T')[0];
}

// Feriados nacionais + estaduais de Sao Paulo + municipais da cidade de Sao Paulo.
// Calculado localmente (sem depender de calendario externo do Google, que so
// tem feriados nacionais) para o Espaco Liquidificador, em Campos Eliseos, SP.
async function getFeriadosDoAno(ano) {
  if (_cacheFeriados[ano]) return _cacheFeriados[ano];

  const feriados = new Set();

  // Nacionais (fixos)
  feriados.add(ano + '-01-01'); // Confraternizacao Universal
  feriados.add(ano + '-04-21'); // Tiradentes
  feriados.add(ano + '-05-01'); // Dia do Trabalho
  feriados.add(ano + '-09-07'); // Independencia do Brasil
  feriados.add(ano + '-10-12'); // Nossa Senhora Aparecida
  feriados.add(ano + '-11-02'); // Finados
  feriados.add(ano + '-11-15'); // Proclamacao da Republica
  feriados.add(ano + '-11-20'); // Dia Nacional de Zumbi e da Consciencia Negra
  feriados.add(ano + '-12-25'); // Natal

  // Estadual (Sao Paulo)
  feriados.add(ano + '-07-09'); // Revolucao Constitucionalista de 1932

  // Municipal (cidade de Sao Paulo)
  feriados.add(ano + '-01-25'); // Aniversario da cidade de Sao Paulo

  // Moveis (calculados a partir da Pascoa)
  const pascoa = calcularPascoa(ano);
  const addDias = (data, dias) => new Date(data.getTime() + dias * 24 * 60 * 60000);
  feriados.add(formatarDataStrUTC(addDias(pascoa, -48))); // Segunda-feira de Carnaval
  feriados.add(formatarDataStrUTC(addDias(pascoa, -47))); // Terca-feira de Carnaval
  feriados.add(formatarDataStrUTC(addDias(pascoa, -2)));  // Sexta-feira Santa
  feriados.add(formatarDataStrUTC(addDias(pascoa, 60)));  // Corpus Christi

  _cacheFeriados[ano] = feriados;
  return feriados;
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
    const deposito = Math.round(totalComNota * 0.3 * 100) / 100;

    res.json({ ok: true, ...resultado, totalFinal: totalComNota, deposito });
  } catch (err) {
    console.error('[sala-ensaio] calcular-ensaio:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ============================================================
// CONTRATO ONLINE — Sala de Ensaio (clickwrap com trilha de auditoria)
// ============================================================
const CONTRATOS_SALA_DB = '9384a7a1-6356-401e-ac48-cb65eb43e68a';
const VERSAO_TEXTO_CONTRATO = '2026-07-v1';

function montarTextoContrato({ nome, rg, cpf, endereco, blocos, valorTotal, tipoEnsaio, aereos, tipoComprovacao, detalhesComprovacao, seguro, contatoEmergenciaNome, contatoEmergenciaTelefone }) {
  const linhasDatas = blocos.map(b => {
    const dataFmt = b.data.split('-').reverse().join('/');
    return dataFmt + ' — ' + b.inicio + ' às ' + b.fim;
  }).join('\n');

  let clausulaAereos = '6.4. Não serão utilizados aparelhos aéreos (vigas, tecido, lira, trapézio, corda) nesta cessão.';
  if (aereos) {
    clausulaAereos = '6.4. O BENEFICIÁRIO declara que utilizará aparelhos aéreos e apresentou comprovação de aptidão na modalidade: ' + tipoComprovacao + '.\nDetalhes: ' + (detalhesComprovacao || '-');
  }

  return 'CONTRATO DE USO DO ESPAÇO LIQUIDIFICADOR\n' +
    'Cessão onerosa e temporária de uso\n\n' +
    '1. DAS PARTES\n\n' +
    'CEDENTE: CRISTIANE SOCCI LEONEL ME, CNPJ 28.398.119/0001-83, nome fantasia Cia. do Liquidificador / Espaço Liquidificador, com sede na Rua Doutor Carvalho de Mendonça, 67, Campos Elíseos, CEP 01201-010, São Paulo/SP.\n\n' +
    'BENEFICIÁRIO(A): ' + nome + ', RG nº ' + rg + ' e CPF nº ' + cpf + ', residente e domiciliado(a) à ' + endereco + '.\n\n' +
    '2. DO OBJETO\n\n' +
    '2.1. A CEDENTE cede ao BENEFICIÁRIO, de forma onerosa e temporária, o uso do espaço de ensaio, exclusivamente para os fins de ' + (tipoEnsaio || 'ensaio') + ', nas datas e horários abaixo.\n' +
    '2.2. A presente cessão é pessoal e intransferível, não gerando ao BENEFICIÁRIO qualquer direito de posse, locação ou vínculo locatício, sendo vedada a sublocação, empréstimo ou cessão a terceiros.\n\n' +
    '3. DAS DATAS E HORÁRIOS DE USO\n\n' + linhasDatas + '\n\n' +
    '4. DO PREÇO E DA FORMA DE PAGAMENTO\n\n' +
    '4.1. Pelo uso do espaço, o BENEFICIÁRIO pagará à CEDENTE o valor total de R$ ' + valorTotal.toFixed(2) + ', apurado conforme a tabela de valores vigente do Espaço Liquidificador.\n' +
    '4.5. O pagamento será integral e antecipado (sinal de 30% para garantir a reserva, saldo até o início do uso), por PIX, com envio do comprovante pelo WhatsApp.\n\n' +
    '5. DA RESPONSABILIDADE POR ACIDENTES\n\n' +
    '5.1. O BENEFICIÁRIO reconhece que as atividades artísticas, físicas e circenses são executadas por sua própria conta, iniciativa e risco, declarando possuir a aptidão técnica necessária.\n' +
    '5.2. O BENEFICIÁRIO responde pelos acidentes e lesões que decorram de sua própria conduta, do uso inadequado do espaço ou equipamentos, ou de atos de seus convidados.\n\n' +
    '6. DOS DIREITOS E DEVERES DO BENEFICIÁRIO\n\n' +
    '6.1. O BENEFICIÁRIO obriga-se a manter o espaço trancado durante ausências e devolver a chave imediatamente ao encerrar as atividades.\n' +
    '6.2. Devolver o espaço limpo e em condições de uso, ressarcindo a CEDENTE por danos causados, ressalvado o desgaste natural.\n' +
    '6.3. É vedado o acesso ao acervo da Cia. do Liquidificador.\n' +
    clausulaAereos + '\n' +
    '6.5. Respeitar a vizinhança, os níveis de ruído e os horários contratados. Vedada a realização de festas ou eventos comerciais sem autorização prévia e por escrito.\n\n' +
    'SEGURO PESSOAL: ' + (seguro || 'Não informado') + '\n' +
    'CONTATO DE EMERGÊNCIA: ' + (contatoEmergenciaNome || '-') + ' — ' + (contatoEmergenciaTelefone || '-') + '\n\n' +
    '9. DA POLÍTICA DE CANCELAMENTO E REMARCAÇÃO\n\n' +
    '9.1. Cancelamento com 48h ou mais de antecedência: direito à remarcação ou devolução do valor.\n' +
    '9.2. Cancelamento entre 24h e 48h: retenção de 50% do valor.\n' +
    '9.3. Cancelamento com menos de 24h ou não comparecimento: retenção integral do valor.\n\n' +
    '10. DA PROTEÇÃO DE DADOS (LGPD)\n\n' +
    '10.1. Os dados pessoais do BENEFICIÁRIO são tratados exclusivamente para a execução deste contrato e cumprimento de obrigações legais (art. 7º, II e V, Lei 13.709/2018), não sendo compartilhados com terceiros salvo exigência legal.\n\n' +
    '12. DA ASSINATURA ELETRÔNICA\n\n' +
    '12.1. As partes reconhecem a validade da assinatura eletrônica deste instrumento, nos termos do art. 10, §2º, da MP 2.200-2/2001.\n\n' +
    '13. DO FORO\n\n' +
    '13.1. Fica eleito o foro da comarca de São Paulo/SP.';
}

async function gerarPdfContrato(textoContrato, { nome, dataHoraISO, ip, dispositivo, assinaturaDigitada }) {
  const PDFDocument = require('pdfkit');
  const chunks = [];
  const doc = new PDFDocument({ margin: 50 });
  doc.on('data', (c) => chunks.push(c));
  const fimPromise = new Promise((resolve) => doc.on('end', () => resolve(Buffer.concat(chunks))));

  doc.fontSize(9).text(textoContrato, { align: 'left' });
  doc.moveDown(2);
  doc.fontSize(9).text('--- ASSINATURA ELETRÔNICA ---');
  doc.text('Assinado por: ' + assinaturaDigitada);
  doc.text('Data/Hora: ' + dataHoraISO);
  doc.text('IP: ' + ip);
  doc.text('Dispositivo: ' + dispositivo);

  doc.end();
  return await fimPromise;
}

async function uploadBufferOneDrive(token, folderId, buffer, filename) {
  const uploadUrl = 'https://graph.microsoft.com/v1.0/users/' + MS_USER + '/drive/items/' + folderId + ':/' + encodeURIComponent(filename) + ':/content';
  const r = await fetch(uploadUrl, {
    method: 'PUT',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/octet-stream' },
    body: buffer,
  });
  if (!r.ok) { const t = await r.text(); throw new Error('Upload OneDrive: ' + t); }
  return await r.json();
}

app.get('/contrato/:reservaId', async (req, res) => {
  const { reservaId } = req.params;
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + SALA_ENSAIO_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Reserva ID', rich_text: { equals: reservaId } }, page_size: 50 }),
    });
    const d = await r.json();
    const registros = d.results || [];
    if (registros.length === 0) return res.json({ ok: false, erro: 'Reserva não encontrada.' });

    const p0 = registros[0].properties;
    const blocos = registros.map(page => {
      const pp = page.properties;
      const inicioISO = pp['Início']?.date?.start || '';
      const fimISO = pp['Fim']?.date?.start || '';
      return {
        data: inicioISO.split('T')[0],
        inicio: inicioISO ? inicioISO.slice(11, 16) : '',
        fim: fimISO ? fimISO.slice(11, 16) : '',
      };
    });

    res.json({
      ok: true,
      projeto: p0['Nome do Projeto']?.rich_text?.[0]?.plain_text || '',
      contatoNome: p0['Contato']?.rich_text?.[0]?.plain_text || '',
      whatsapp: p0['WhatsApp']?.rich_text?.[0]?.plain_text || '',
      tipoEnsaio: p0['Tipo de Ensaio']?.select?.name || '',
      valorTotal: p0['Valor Total']?.number || 0,
      status: p0['Status']?.select?.name || '',
      blocos,
    });
  } catch (err) {
    console.error('[contrato] erro ao buscar reserva:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/contrato/:reservaId/aceitar', async (req, res) => {
  const { reservaId } = req.params;
  const {
    nome, rg, cpf, endereco, aereos, tipoComprovacao, detalhesComprovacao, comprovanteArquivo, comprovanteNomeArquivo,
    seguro, contatoEmergenciaNome, contatoEmergenciaTelefone, aceite, assinaturaDigitada, dispositivo,
  } = req.body;
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();

  if (!nome || !rg || !cpf || !endereco || !aceite || !assinaturaDigitada || !contatoEmergenciaNome || !contatoEmergenciaTelefone) {
    return res.status(400).json({ ok: false, erro: 'Preencha todos os campos obrigatórios.' });
  }
  if (aereos && (!tipoComprovacao || tipoComprovacao === '' || (tipoComprovacao !== 'Autodeclaração' && !comprovanteArquivo))) {
    return res.status(400).json({ ok: false, erro: 'Preencha a comprovação de aptidão para aéreos.' });
  }

  try {
    const rReserva = await fetch('https://api.notion.com/v1/databases/' + SALA_ENSAIO_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { property: 'Reserva ID', rich_text: { equals: reservaId } }, page_size: 50 }),
    });
    const dReserva = await rReserva.json();
    const registros = dReserva.results || [];
    if (registros.length === 0) return res.json({ ok: false, erro: 'Reserva não encontrada.' });

    const p0 = registros[0].properties;
    const blocos = registros.map(page => {
      const pp = page.properties;
      const inicioISO = pp['Início']?.date?.start || '';
      const fimISO = pp['Fim']?.date?.start || '';
      return { data: inicioISO.split('T')[0], inicio: inicioISO.slice(11, 16), fim: fimISO.slice(11, 16) };
    });
    const valorTotal = p0['Valor Total']?.number || 0;
    const tipoEnsaio = p0['Tipo de Ensaio']?.select?.name || '';

    const msToken = await getMicrosoftToken();
    const nomePasta = 'Contratos-SalaEnsaio-' + reservaId;
    const folderId = await criarOuObterSubpasta(msToken, nomePasta);

    let linkComprovante = '';
    if (comprovanteArquivo) {
      const base64Clean = comprovanteArquivo.replace(/^data:.*;base64,/, '');
      const buffer = Buffer.from(base64Clean, 'base64');
      const nomeArquivo = comprovanteNomeArquivo || 'comprovante-aptidao';
      const uploaded = await uploadBufferOneDrive(msToken, folderId, buffer, nomeArquivo);
      linkComprovante = await criarLinkCompartilhamento(msToken, uploaded.id);
    }

    const dataHoraISO = new Date().toISOString();
    const textoContrato = montarTextoContrato({
      nome, rg, cpf, endereco, blocos, valorTotal, tipoEnsaio,
      aereos, tipoComprovacao, detalhesComprovacao, seguro, contatoEmergenciaNome, contatoEmergenciaTelefone,
    });

    const pdfBuffer = await gerarPdfContrato(textoContrato, { nome, dataHoraISO, ip, dispositivo, assinaturaDigitada });
    const pdfUploaded = await uploadBufferOneDrive(msToken, folderId, pdfBuffer, 'contrato-assinado-' + reservaId + '.pdf');
    const linkPdf = await criarLinkCompartilhamento(msToken, pdfUploaded.id);

    await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        parent: { database_id: CONTRATOS_SALA_DB },
        properties: {
          'Título': { title: [{ text: { content: nome + ' — ' + reservaId } }] },
          'Reserva ID': { rich_text: [{ text: { content: reservaId } }] },
          'Nome Completo': { rich_text: [{ text: { content: nome } }] },
          'RG': { rich_text: [{ text: { content: rg } }] },
          'CPF': { rich_text: [{ text: { content: cpf.replace(/\D/g, '') } }] },
          'Endereço': { rich_text: [{ text: { content: endereco } }] },
          'Envolve Aéreos': { select: { name: aereos ? 'Sim' : 'Não' } },
          'Tipo de Comprovação': { select: { name: aereos ? tipoComprovacao : 'Não aplicável' } },
          'Detalhes da Comprovação': { rich_text: [{ text: { content: detalhesComprovacao || '' } }] },
          'Link do Comprovante': { url: linkComprovante || null },
          'Tem Seguro Pessoal': { select: { name: seguro || 'Não informado' } },
          'Contato de Emergência': { rich_text: [{ text: { content: contatoEmergenciaNome } }] },
          'Telefone de Emergência': { phone_number: contatoEmergenciaTelefone },
          'Aceite Confirmado': { checkbox: true },
          'Assinatura Digitada': { rich_text: [{ text: { content: assinaturaDigitada } }] },
          'Data/Hora do Aceite': { date: { start: dataHoraISO } },
          'IP': { rich_text: [{ text: { content: ip || '' } }] },
          'Dispositivo': { rich_text: [{ text: { content: dispositivo || '' } }] },
          'Versão do Texto': { rich_text: [{ text: { content: VERSAO_TEXTO_CONTRATO } }] },
          'Link do PDF Gerado': { url: linkPdf || null },
          'Status': { select: { name: 'Aceito' } },
        },
      }),
    });

    const primeiroNome = nome.split(' ')[0];
    const numLimpo = (p0['WhatsApp']?.rich_text?.[0]?.plain_text || '').replace(/\D/g, '');
    const numBr = numLimpo.length === 11 ? '55' + numLimpo : numLimpo;
    if (numBr) {
      try { await enviarWhatsApp(numBr, 'Prontinho, ' + primeiroNome + '! ✅\n\nSeu contrato foi assinado com sucesso. Aqui está o PDF:\n' + linkPdf); } catch(e) {}
    }
    const msgInterna = '📜 Contrato assinado — Reserva ' + reservaId + ' (' + nome + ')\nPDF: ' + linkPdf;
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}

    res.json({ ok: true, linkPdf });
  } catch (err) {
    console.error('[contrato] erro ao aceitar:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/verificar-conflito', async (req, res) => {
  const { data, inicio, fim } = req.body;
  if (!data || !inicio || !fim) return res.status(400).json({ error: 'data, inicio, fim obrigatorios' });

  try {
    const inicioISO = data + 'T' + inicio + ':00-03:00';
    const fimISO = data + 'T' + fim + ':00-03:00';
    const resultado = await verificarDisponibilidade(inicioISO, fimISO);
    res.json({ ok: true, disponivel: resultado.disponivel, ehResidente: !!resultado.ehResidente, semana: getNumeroSemanaISO(data) });
  } catch (err) {
    console.error('[sala-ensaio] verificar-conflito:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get('/residente-dia/:data', async (req, res) => {
  const data = req.params.data;
  if (!data) return res.status(400).json({ ok: false, error: 'data obrigatoria' });
  try {
    // Janela do dia inteiro (07h as 22h, horario de funcionamento da sala)
    const timeMinISO = data + 'T07:00:00-03:00';
    const timeMaxISO = data + 'T22:00:00-03:00';
    const intervalosRaw = await obterOcupacaoResidente(timeMinISO, timeMaxISO);
    const intervalosComBuffer = aplicarBufferIntervalos(intervalosRaw);
    const intervalos = intervalosComBuffer.map(iv => {
      const inicioBr = new Date(iv.start.getTime() - 3 * 60 * 60000);
      const fimBr = new Date(iv.end.getTime() - 3 * 60 * 60000);
      return {
        inicio: inicioBr.toISOString().split('T')[1].slice(0, 5),
        fim: fimBr.toISOString().split('T')[1].slice(0, 5),
      };
    });
    res.json({ ok: true, ehResidente: intervalos.length > 0, intervalos, semana: getNumeroSemanaISO(data) });
  } catch (err) {
    console.error('[residente-dia] erro:', err.message);
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
    const deposito = Math.round(totalFinal * 0.3 * 100) / 100;

    const blocosResidentes = [];
    const semanasResidentesNesteRequest = new Set();
    let totalResidentesNesteRequest = 0;
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
      if (resultadoDisp.ehResidente) {
        // Checagem extra: dentro desta MESMA reserva (varios dias de uma vez),
        // nao deixa usar 2+ dias da Cia Pla na mesma semana nem mais de 4 no total,
        // ja que o registro no Notion so acontece depois que o loop termina.
        const semanaBloco = getNumeroSemanaISO(b.data);
        if (semanasResidentesNesteRequest.has(semanaBloco) || totalResidentesNesteRequest >= 4) {
          const motivo = 'O dia ' + b.data + ' das ' + b.inicio + ' às ' + b.fim + ' também pertence à residência da Cia Plá, e esta reserva já está usando outro horário deles na mesma semana/mês. Escolha outro horário para um dos dias.';
          return res.json({ ok: false, erro: motivo });
        }
        semanasResidentesNesteRequest.add(semanaBloco);
        totalResidentesNesteRequest++;
        blocosResidentes.push(b);
      }
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

    const notaFiscalTextoInterno = notaFiscal ? '\n📄 Nota fiscal solicitada — CPF/CNPJ: ' + (cpfCnpj || 'A informar') : '';
    const msgInterna = '🎭 *Novo pré-agendamento — Sala de Ensaio*\n\nProjeto: ' + projeto + '\nColetivo: ' + (coletivo || '-') + '\nResponsável: ' + (diretor || '-') + '\nContato: ' + (contatoNome || '-') + ' (' + whatsapp + ')\nPessoas: ' + qtdPessoas + '\n' + resumoDias + '\n\nValor: R$ ' + totalFinal.toFixed(2) + ' (sinal R$ ' + deposito.toFixed(2) + ')' + notaFiscalTextoInterno + '\nReserva: ' + reservaId;
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) { console.error('[sala-ensaio] wpp fabio:', e.message); }
    try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) { console.error('[sala-ensaio] wpp cia:', e.message); }

    res.json({ ok: true, reservaId, totalFinal, deposito });
  } catch (err) {
    console.error('[sala-ensaio] reservar-sala:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

async function expirarOfertaResidenteSeExistir(dataStr) {
  try {
    const r = await fetch('https://api.notion.com/v1/databases/' + RESIDENTES_REMARCACOES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ filter: { and: [
        { property: 'Data Original', date: { equals: dataStr } },
        { property: 'Status', select: { equals: 'Oferecida' } },
      ]}, page_size: 10 }),
    });
    const d = await r.json();
    for (const page of (d.results || [])) {
      await fetch('https://api.notion.com/v1/pages/' + page.id, {
        method: 'PATCH',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ properties: { 'Status': { select: { name: 'Expirada' } } } }),
      });
      console.log('[residentes] oferta de remarcacao expirada para ' + dataStr + ' (reserva cancelada)');
    }
  } catch (e) {
    console.error('[residentes] erro ao expirar oferta:', e.message);
  }
}

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

      const inicioStr = page.properties?.['Início']?.date?.start || '';
      if (inicioStr) {
        const dataStr = inicioStr.split('T')[0];
        await expirarOfertaResidenteSeExistir(dataStr);
      }
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
      const linkContrato = 'https://agende-ensaio.ciadoliquidificador.com.br/contrato/?reserva=' + estado.reservaId;
      await enviarWhatsApp(numero, 'Perfeito, ' + estado.nome + '! ✅\n\nAguardamos o comprovante do sinal (R$ ' + estado.deposito.toFixed(2) + ') por aqui — pode mandar a foto ou print do PIX.\n\nTambém precisamos que você preencha e assine o contrato de uso do espaço, é rapidinho:\n' + linkContrato);
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
async function obterOcupacaoResidente(timeMinISO, timeMaxISO) {
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/calendar.readonly'],
    });
    const authClient = await auth.getClient();
    const calendar = google.calendar({ version: 'v3', auth: authClient });
    const resp = await calendar.freebusy.query({
      requestBody: { timeMin: timeMinISO, timeMax: timeMaxISO, items: [{ id: RESIDENTE_CIA_PLA_CALENDAR }] },
    });
    const busy = resp.data.calendars?.[RESIDENTE_CIA_PLA_CALENDAR]?.busy || [];
    return busy.map(b => ({ start: new Date(b.start), end: new Date(b.end) }));
  } catch (e) {
    console.error('[disponibilidade-mes] erro ao buscar ocupacao residente:', e.message);
    return [];
  }
}

app.get('/disponibilidade-mes', async (req, res) => {
  const { ano, mes } = req.query;
  if (!ano || !mes) return res.status(400).json({ error: 'ano e mes obrigatorios' });

  try {
    const anoNum = parseInt(ano, 10);
    const mesNum = parseInt(mes, 10);
    const timeMinISO = new Date(Date.UTC(anoNum, mesNum - 1, 1, 3, 0, 0)).toISOString();
    const timeMaxISO = new Date(Date.UTC(anoNum, mesNum, 1, 2, 59, 59)).toISOString();

    const [ocupGoogle, ocupNotion, ocupResidenteBruto] = await Promise.all([
      obterOcupacaoGoogle(timeMinISO, timeMaxISO),
      obterOcupacaoNotion(timeMinISO, timeMaxISO),
      obterOcupacaoResidente(timeMinISO, timeMaxISO),
    ]);

    // So mostramos o horario da Cia Pla como ocupado nos dias em que a cota
    // mensal de remarcacoes deles ja estiver esgotada (mesma regra usada na
    // hora de confirmar a reserva). Enquanto tiver cota, o dia fica livre no
    // calendario visual, porque o cliente ainda pode negociar aquele horario.
    const diasResidenteEsgotados = new Set();
    for (const iv of ocupResidenteBruto) {
      const diaStr = new Date(iv.start.getTime() - 3 * 60 * 60000).toISOString().split('T')[0];
      if (diasResidenteEsgotados.has(diaStr)) continue;
      const podeUsar = await residenteDisponivelParaTerceiro(diaStr);
      if (!podeUsar) diasResidenteEsgotados.add(diaStr);
    }
    const ocupResidenteFiltrada = ocupResidenteBruto.filter(iv => {
      const diaStr = new Date(iv.start.getTime() - 3 * 60 * 60000).toISOString().split('T')[0];
      return diasResidenteEsgotados.has(diaStr);
    });

    const todasOcupacoes = aplicarBufferIntervalos([...ocupGoogle, ...ocupNotion, ...ocupResidenteFiltrada]);

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

// ============================================================
// MANUTENCAO — Limpar eventos orfaos no calendario Ensaio Externo
// (eventos cujo registro no Notion foi apagado ou marcado Cancelado
// direto na interface, sem passar pelo fluxo automatico)
// ============================================================
app.get('/limpar-eventos-orfaos', async (req, res) => {
  try {
    const calendar = await getGoogleCalendarClient();
    const agora = new Date();
    const timeMin = new Date(agora.getTime() - 30 * 24 * 60 * 60000).toISOString();
    const timeMax = new Date(agora.getTime() + 365 * 24 * 60 * 60000).toISOString();

    const listaResp = await calendar.events.list({
      calendarId: CALENDARIO_ENSAIO_EXTERNO,
      timeMin, timeMax,
      singleEvents: true,
      maxResults: 2500,
    });
    const eventos = listaResp.data.items || [];

    const removidos = [];
    const mantidos = [];
    const semReservaId = [];

    for (const ev of eventos) {
      const desc = ev.description || '';
      const match = desc.match(/Reserva:\s*(ENS-\d+)/);
      if (!match) { semReservaId.push({ id: ev.id, summary: ev.summary, start: ev.start }); continue; }
      const reservaId = match[1];

      const r = await fetch('https://api.notion.com/v1/databases/' + SALA_ENSAIO_DB + '/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ filter: { property: 'Reserva ID', rich_text: { equals: reservaId } }, page_size: 5 }),
      });
      const data = await r.json();
      const registroAtivo = (data.results || []).some(p => p.properties?.Status?.select?.name !== 'Cancelado');

      if (registroAtivo) {
        mantidos.push({ reservaId, start: ev.start });
      } else {
        try {
          await calendar.events.delete({ calendarId: CALENDARIO_ENSAIO_EXTERNO, eventId: ev.id });
          removidos.push({ reservaId, start: ev.start, summary: ev.summary });
        } catch (e) {
          console.error('[limpar-eventos-orfaos] erro ao excluir evento ' + ev.id + ':', e.message);
        }
      }
    }

    res.json({
      ok: true,
      totalEventosVerificados: eventos.length,
      removidos,
      mantidos: mantidos.length,
      semReservaId,
    });
  } catch (err) {
    console.error('[limpar-eventos-orfaos] erro:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});


// Verifica a cada 15 minutos se algum pre-agendamento de sala ficou sem resposta por 3h+ (em horario comercial)
setInterval(async () => {
  const agora = Date.now();
  const dentroHorarioComercial = (await calcularProximoHorarioComercial()) === null;
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
// WEBHOOK — Cancelamento disparado por Automation do Notion
// (quando alguem muda o Status para "Cancelado" direto na tabela)
// ============================================================
app.post('/webhook-cancelamento-notion', async (req, res) => {
  res.status(200).json({ ok: true }); // responde rapido, processa depois

  try {
    const body = req.body || {};
    console.log('[webhook-cancelamento-notion] payload recebido:', JSON.stringify(body).slice(0, 1000));

    const pageId = (body.data && body.data.id) || body.pageId || body.page_id || null;
    if (!pageId) {
      console.error('[webhook-cancelamento-notion] payload sem page id reconhecivel.');
      return;
    }

    // Sempre busca a pagina completa e atualizada pelo ID, em vez de confiar
    // nas propriedades que a automation do Notion decidiu incluir no payload
    // (essas variam conforme o que foi marcado na configuracao da automation).
    const rPage = await fetch('https://api.notion.com/v1/pages/' + pageId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const pageData = await rPage.json();

    const status = pageData.properties?.Status?.select?.name || '';
    if (status !== 'Cancelado') {
      console.log('[webhook-cancelamento-notion] status atual da pagina nao e Cancelado (' + status + '), ignorando.');
      return;
    }

    const eventId = pageData.properties?.['Google Event ID']?.rich_text?.[0]?.plain_text || '';
    if (eventId) {
      await excluirEventoEnsaioExterno(eventId);
      console.log('[webhook-cancelamento-notion] evento removido do Calendar: ' + eventId);
    } else {
      console.log('[webhook-cancelamento-notion] pagina sem Google Event ID, nada para remover do Calendar.');
    }

    const inicioStr = pageData.properties?.['Início']?.date?.start || '';
    if (inicioStr) {
      const dataStr = inicioStr.split('T')[0];
      await expirarOfertaResidenteSeExistir(dataStr);
    }
  } catch (err) {
    console.error('[webhook-cancelamento-notion] erro:', err.message);
  }
});

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
  const resp = await fetch('https://api.notion.com/v1/pages', {
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
  const data = await resp.json();
  return data.id;
}

// ============================================================
// REMARCACAO DIRETA DO RESIDENTE (Cia Pla) - sem fluxo de pagamento
// ============================================================
app.get('/remarcacao/:id', async (req, res) => {
  try {
    const r = await fetch('https://api.notion.com/v1/pages/' + req.params.id, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    if (!r.ok) return res.json({ ok: false, erro: 'Oferta nao encontrada.' });
    const page = await r.json();
    const p = page.properties;
    const status = p['Status']?.select?.name || '';
    const residente = p['Residente']?.select?.name || '';
    const dataOriginal = p['Data Original']?.date?.start || '';
    const dataRemarcada = p['Data Remarcada']?.date?.start || '';
    res.json({ ok: true, status, residente, dataOriginal, dataRemarcada });
  } catch (err) {
    console.error('[remarcacao] erro ao buscar oferta:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

app.post('/confirmar-remarcacao', async (req, res) => {
  const { ofertaId, data, inicio, fim } = req.body;
  if (!ofertaId || !data || !inicio || !fim) return res.status(400).json({ ok: false, erro: 'Campos obrigatorios faltando.' });

  try {
    const rOferta = await fetch('https://api.notion.com/v1/pages/' + ofertaId, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    if (!rOferta.ok) return res.json({ ok: false, erro: 'Oferta nao encontrada.' });
    const oferta = await rOferta.json();
    const statusAtual = oferta.properties['Status']?.select?.name || '';
    if (statusAtual !== 'Oferecida') {
      return res.json({ ok: false, erro: 'Essa remarcacao ja foi usada ou expirou. Fale com a equipe do espaco.' });
    }

    const inicioISO = data + 'T' + inicio + ':00-03:00';
    const fimISO = data + 'T' + fim + ':00-03:00';

    const disponibilidade = await verificarDisponibilidade(inicioISO, fimISO);
    if (!disponibilidade.disponivel) {
      return res.json({ ok: false, erro: 'Esse horario ja esta ocupado. Escolha outro.' });
    }

    const calendar = await getGoogleCalendarClient();
    await calendar.events.insert({
      calendarId: RESIDENTE_CIA_PLA_CALENDAR,
      requestBody: {
        summary: 'Cia Plá — Ensaio remarcado',
        description: 'Remarcacao referente a oferta ' + ofertaId,
        start: { dateTime: inicioISO },
        end: { dateTime: fimISO },
      },
    });

    await fetch('https://api.notion.com/v1/pages/' + ofertaId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties: {
        'Status': { select: { name: 'Remarcada' } },
        'Data Remarcada': { date: { start: inicioISO } },
      }}),
    });

    const dataFmt = data.split('-').reverse().join('/');
    const msgConfirmacao = 'Prontinho! ✅\n\nSeu ensaio foi remarcado para ' + dataFmt + ', das ' + inicio + ' às ' + fim + '.';
    try { await enviarWhatsAppComHorarioComercial(WHATSAPP_CIA_PLA, msgConfirmacao); } catch(e) {}
    const msgInterna = '🔄 Cia Plá remarcou o ensaio para ' + dataFmt + ' (' + inicio + '-' + fim + ') — oferta ' + ofertaId;
    try { await enviarWhatsApp(WHATSAPP_FABIO, msgInterna); } catch(e) {}
    try { await enviarWhatsApp(WHATSAPP_CIA, msgInterna); } catch(e) {}

    res.json({ ok: true });
  } catch (err) {
    console.error('[confirmar-remarcacao] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

async function notificarResidenteSobreposicao(blocosAfetados) {
  for (const b of blocosAfetados) {
    const dataFmt = b.data.split('-').reverse().join('/');
    const semanaStr = getNumeroSemanaISO(b.data);
    const mesStr = b.data.slice(0, 7);

    const { totalMes, naMesmaSemana } = await contarRemarcacoesResidente('Cia Plá', mesStr, semanaStr);

    const ofertaId = await registrarOfertaRemarcacao('Cia Plá', b.data);

    const limiteAtingido = totalMes >= 4 || naMesmaSemana >= 1;
    const avisoLimite = limiteAtingido
      ? '\n\n⚠️ Atenção: o limite de remarcações do mês (4x, sendo só 1x por semana) já foi atingido ou está no limite. Fale com a gente para verificar.'
      : '\n\nVocês já usaram ' + totalMes + ' de 4 remarcações este mês.';

    const linkRemarcacao = 'https://remarcar-residente.ciadoliquidificador.com.br?oferta=' + ofertaId;
    const msg = 'Olá! 🎭\n\nO horário de ensaio de vocês do dia ' + dataFmt + ' (' + b.inicio + ' às ' + b.fim + ') foi reservado por um cliente pagante, conforme nosso acordo.\n\nVocês podem remarcar essas horas direto por aqui, sem passar por pagamento:\n' + linkRemarcacao + avisoLimite;

    try {
      await enviarWhatsAppComHorarioComercial(WHATSAPP_CIA_PLA, msg);
    } catch (e) {
      console.error('[residentes] erro ao notificar Cia Pla:', e.message);
    }
  }
}

// ============================================================
// CALCULADORA DE ORÇAMENTO — integração com Notion (2026 - PROPOSTAS E CONTRATOS)
// ============================================================
const ORCAMENTO_PROPOSTAS_DB = '2c6c45031f73804f8f90e6e7439d7e1c';
const ORCAMENTO_TRABALHOS_DB = '4589d769656b41149e9bf6300b30d886';

// GET /orcamento/opcoes-notion — listas reais para alimentar os dropdowns da calculadora
app.get('/orcamento/opcoes-notion', async (req, res) => {
  try {
    const rTrabalhos = await fetch('https://api.notion.com/v1/databases/' + ORCAMENTO_TRABALHOS_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ page_size: 100, sorts: [{ property: 'Nome', direction: 'ascending' }] }),
    });
    const dTrabalhos = await rTrabalhos.json();
    const trabalhos = (dTrabalhos.results || []).map(p => ({
      id: p.id,
      nome: p.properties?.Nome?.title?.[0]?.plain_text || '(sem nome)',
    }));

    const rIntegrantes = await fetch('https://api.notion.com/v1/databases/' + INTEGRANTES_DB + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ page_size: 100, sorts: [{ property: 'Nome', direction: 'ascending' }] }),
    });
    const dIntegrantes = await rIntegrantes.json();
    const integrantes = (dIntegrantes.results || []).map(p => ({
      id: p.id,
      nome: p.properties?.Nome?.title?.[0]?.plain_text || '(sem nome)',
    }));

    const rSchema = await fetch('https://api.notion.com/v1/databases/' + ORCAMENTO_PROPOSTAS_DB, {
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28' },
    });
    const dSchema = await rSchema.json();
    const contratanteOpcoes = (dSchema.properties?.['Contratante']?.multi_select?.options || [])
      .map(o => o.name)
      .sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const tipoOpcoes = (dSchema.properties?.['Tipo']?.select?.options || []).map(o => o.name);

    res.json({ ok: true, trabalhos, integrantes, contratante: contratanteOpcoes, tipo: tipoOpcoes });
  } catch (err) {
    console.error('[orcamento/opcoes-notion] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});

// POST /orcamento/salvar-notion — cria as páginas da apresentação (uma por data/horário),
// gera a planilha .xlsx, sobe pro OneDrive e grava o link em cada página criada.
app.post('/orcamento/salvar-notion', async (req, res) => {
  const {
    local, endereco, contratante, tipo, trabalhoId, integrantesIds, contato,
    apresentacoes, valorPorApresentacao, cacheElenco, cacheProducao, cacheTecnicos,
    detalhamento,
  } = req.body;

  if (!local || !tipo || !contratante || !Array.isArray(apresentacoes) || apresentacoes.length === 0) {
    return res.status(400).json({ ok: false, erro: 'Preencha local, contratante, tipo e ao menos uma data/horário.' });
  }
  if (valorPorApresentacao === undefined || valorPorApresentacao === null) {
    return res.status(400).json({ ok: false, erro: 'Valor por apresentação é obrigatório.' });
  }

  try {
    const XLSX = require('xlsx');
    const linhasResumo = [
      ['Orçamento — ' + local],
      [],
      ['Contratante', contratante],
      ['Tipo', tipo],
      ['Contato', contato || ''],
      [],
      ['Datas e horários'],
      ...apresentacoes.map(a => [a.data || '', a.horario || '']),
      [],
      ['Valor por apresentação', valorPorApresentacao],
      ['Quantidade de apresentações', apresentacoes.length],
      ['Valor total', Math.round(valorPorApresentacao * apresentacoes.length * 100) / 100],
      ['Cachê Elenco', cacheElenco || 0],
      ['Cachê Produção', cacheProducao || 0],
      ['Cachê Técnicos', cacheTecnicos || 0],
    ];
    if (detalhamento && detalhamento.percentuais && detalhamento.percentuais.length) {
      linhasResumo.push([], ['Percentuais'], ['Nome', '%']);
      detalhamento.percentuais.forEach(p => linhasResumo.push([p.nome, p.valor]));
    }
    if (detalhamento && detalhamento.custosPessoais && detalhamento.custosPessoais.length) {
      linhasResumo.push([], ['Custos Pessoais'], ['Categoria', 'Qtd. pessoas', 'Valor/pessoa']);
      detalhamento.custosPessoais.forEach(c => linhasResumo.push([c.nome, c.qtd, c.valor]));
    }
    if (detalhamento && detalhamento.custosMateriais && detalhamento.custosMateriais.length) {
      linhasResumo.push([], ['Custos Materiais'], ['Item', 'Valor']);
      detalhamento.custosMateriais.forEach(m => linhasResumo.push([m.nome, m.valor]));
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(linhasResumo);
    XLSX.utils.book_append_sheet(wb, ws, 'Orçamento');
    const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    let linkPlanilha = '';
    try {
      const msToken = await getMicrosoftToken();
      const nomePasta = 'Orcamentos-Propostas-' + slugify(local);
      const folderId = await criarOuObterSubpasta(msToken, nomePasta);
      const uploaded = await uploadBufferOneDrive(msToken, folderId, xlsxBuffer, 'orcamento-' + Date.now() + '.xlsx');
      linkPlanilha = await criarLinkCompartilhamento(msToken, uploaded.id);
    } catch (eXlsx) {
      console.error('[orcamento/salvar-notion] erro ao gerar/subir planilha:', eXlsx.message);
    }

    // Geocodifica o endereço (necessário para o campo "place" do Notion, que exige
    // latitude/longitude — confirmado ao vivo via teste). Melhor esforço: se falhar,
    // o campo Endereço simplesmente não é gravado (fallback já existente).
    let coordenadasEndereco = null;
    if (endereco && GOOGLE_ROUTES_API_KEY) {
      try {
        const rGeocode = await fetch('https://maps.googleapis.com/maps/api/geocode/json?address=' + encodeURIComponent(endereco) + '&key=' + GOOGLE_ROUTES_API_KEY);
        const dGeocode = await rGeocode.json();
        const local0 = dGeocode.results?.[0]?.geometry?.location;
        if (local0) {
          coordenadasEndereco = { latitude: local0.lat, longitude: local0.lng };
        } else {
          console.error('[orcamento/salvar-notion] geocoding sem resultado para endereço:', endereco, JSON.stringify(dGeocode).slice(0, 300));
        }
      } catch (eGeo) {
        console.error('[orcamento/salvar-notion] erro ao geocodificar endereço:', eGeo.message);
      }
    }

    const paginasCriadas = [];
    let enderecoGravadoEmAlgumaPagina = false;

    for (const apresentacao of apresentacoes) {
      const propsBase = {
        'Local': { title: [{ text: { content: local } }] },
        'Tipo': { select: { name: tipo } },
        'Contratante': { multi_select: [{ name: contratante }] },
        'Contato': { rich_text: [{ text: { content: contato || '' } }] },
        'Data': { date: { start: apresentacao.data } },
        'Horário Apresentação': { rich_text: [{ text: { content: apresentacao.horario || '' } }] },
        'Valor': { number: Number(valorPorApresentacao) },
        'Status': { status: { name: 'Orçamento' } },
      };
      if (cacheElenco !== undefined) propsBase['Cachê Elenco'] = { number: Number(cacheElenco) || 0 };
      if (cacheProducao !== undefined) propsBase['Cachê Produção'] = { number: Number(cacheProducao) || 0 };
      if (cacheTecnicos !== undefined) propsBase['Cachê Técnicos'] = { number: Number(cacheTecnicos) || 0 };
      if (trabalhoId) propsBase['🎭 Trabalhos'] = { relation: [{ id: trabalhoId }] };
      if (Array.isArray(integrantesIds) && integrantesIds.length) {
        propsBase['Integrantes'] = { relation: integrantesIds.map(id => ({ id })) };
      }
      if (linkPlanilha) propsBase['Link da Planilha'] = { url: linkPlanilha };

      const rCriar = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
        body: JSON.stringify({ parent: { database_id: ORCAMENTO_PROPOSTAS_DB }, properties: propsBase }),
      });
      if (!rCriar.ok) {
        const eBody = await rCriar.text();
        console.error('[orcamento/salvar-notion] Notion recusou criar página para ' + apresentacao.data + ':', eBody);
        continue;
      }
      const paginaCriada = await rCriar.json();
      paginasCriadas.push({ id: paginaCriada.id, url: paginaCriada.url, data: apresentacao.data });

      if (endereco && coordenadasEndereco) {
        try {
          const rEndereco = await fetch('https://api.notion.com/v1/pages/' + paginaCriada.id, {
            method: 'PATCH',
            headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
            body: JSON.stringify({
              properties: {
                'Endereço': {
                  place: {
                    name: local,
                    address: endereco,
                    lat: coordenadasEndereco.latitude,
                    lng: coordenadasEndereco.longitude,
                  },
                },
              },
            }),
          });
          if (rEndereco.ok) {
            enderecoGravadoEmAlgumaPagina = true;
          } else {
            const eBodyEnd = await rEndereco.text();
            console.error('[orcamento/salvar-notion] falha ao gravar Endereço (place) na página ' + paginaCriada.id + ':', eBodyEnd);
          }
        } catch (eEnd) {
          console.error('[orcamento/salvar-notion] erro ao tentar gravar Endereço:', eEnd.message);
        }
      }
    }

    if (paginasCriadas.length === 0) {
      return res.status(500).json({ ok: false, erro: 'Nenhuma página foi criada — ver logs para detalhes.' });
    }

    try {
      const linhasDatas = apresentacoes.map(a => '📅 ' + a.data + ' — ' + (a.horario || '')).join('\n');
      const avisoEndereco = endereco ? ('\n📍 Endereço: ' + endereco + (enderecoGravadoEmAlgumaPagina ? '' : ' (não gravou automaticamente no campo — confira/preencha manualmente)')) : '';
      const msg = '💰 *Orçamento salvo* — ' + local + '\n\nContratante: ' + contratante + '\nTipo: ' + tipo + '\n' + linhasDatas +
        '\n\nValor por apresentação: R$ ' + Number(valorPorApresentacao).toFixed(2) + avisoEndereco +
        (linkPlanilha ? ('\n📄 Planilha: ' + linkPlanilha) : '');
      await enviarWhatsApp(WHATSAPP_FABIO, msg);
    } catch (eNotif) {
      console.error('[orcamento/salvar-notion] erro ao notificar Fábio:', eNotif.message);
    }

    res.json({ ok: true, paginas: paginasCriadas, linkPlanilha, enderecoGravado: enderecoGravadoEmAlgumaPagina });
  } catch (err) {
    console.error('[orcamento/salvar-notion] erro:', err.message);
    res.status(500).json({ ok: false, erro: err.message });
  }
});
