const express = require('express');
const fetch = require('node-fetch');

const app = express();

app.use(express.json());

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.post('/', async (req, res) => {
  const { numero, texto } = req.body;

  if (!numero || !texto) {
    return res.json({ ok: false, erro: 'Campos "numero" e "texto" são obrigatórios.' });
  }

  try {
    const response = await fetch('https://ciadoliquidificador.digisac.biz/api/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer c2e0a3ac3ae20585924b2dbc133d68d779770199',
      },
      body: JSON.stringify({
        text: texto,
        type: 'chat',
        serviceId: '012587f9-21ea-4143-9005-c0fbdf109f05',
        number: numero,
        userId: 'b0bb99db-a668-403a-af70-efc1d4a7259a',
        origin: 'bot',
      }),
    });

    if (!response.ok) {
      const text = await response.text();
      return res.json({ ok: false, erro: `Digisac retornou ${response.status}: ${text}` });
    }

    return res.json({ ok: true });
  } catch (err) {
    return res.json({ ok: false, erro: err.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Proxy rodando na porta ${PORT}`);
});
