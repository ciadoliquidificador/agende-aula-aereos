// Só consulta (nunca cria) se os bancos de orçamento de um ano existem no Notion.
// Uso: railway run node .claude/skills/virada-de-ano/checar-bancos.js 2027
const ano = process.argv[2];
if (!/^\d{4}$/.test(ano || '')) { console.error('Uso: railway run node checar-bancos.js <ano>'); process.exit(1); }

const titulos = [ano + ' - PROPOSTAS E CONTRATOS', 'APRESENTAÇÕES ' + ano];

(async () => {
  const r = await fetch('https://api.notion.com/v1/search', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + process.env.NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    body: JSON.stringify({ query: ano, filter: { property: 'object', value: 'database' }, page_size: 100 }),
  });
  if (!r.ok) { console.error('Notion ' + r.status + ': ' + (await r.text()).slice(0, 300)); process.exit(1); }
  const d = await r.json();
  for (const titulo of titulos) {
    const db = (d.results || []).find(x => (x.title || []).map(t => t.plain_text).join('') === titulo);
    console.log((db ? '✅ ' : '❌ ') + titulo + (db ? '  ' + db.url : '  (não existe ainda)'));
  }
})();
