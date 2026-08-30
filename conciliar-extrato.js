// Script de uso único/periódico: concilia um extrato bancário (CSV Nubank) com os pagamentos
// pendentes no banco "Pagamentos" do Notion, e marca como "Pago" os que baterem nome + valor.
//
// Uso:
//   railway run node conciliar-extrato.js <arquivo.csv> --dry-run
//   railway run node conciliar-extrato.js <arquivo.csv> --confirm

const fs = require('fs');

const NOTION_TOKEN = process.env.NOTION_TOKEN;
const PAGAMENTOS_DATA_SOURCE_ID = 'e7e26a17-ee03-4d0f-a9e5-82dbbb38400b';

const args = process.argv.slice(2);
const csvPath = args.find(a => !a.startsWith('--'));
const DRY_RUN = !args.includes('--confirm');

if (!csvPath) {
  console.error('Uso: node conciliar-extrato.js <arquivo.csv> [--confirm]');
  process.exit(1);
}

function normalizar(str) {
  return String(str || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

// ---------------------------------------------------------------------------
// 1. Leitura do extrato (CSV Nubank: Data,Valor,Identificador,Descrição)
// ---------------------------------------------------------------------------

function parseCsvLine(line) {
  // Descrição pode teoricamente conter vírgulas, então captura só as 3 primeiras colunas
  // e deixa o resto (tudo depois da 3ª vírgula) como Descrição.
  const m = line.match(/^([^,]*),([^,]*),([^,]*),(.*)$/);
  if (!m) return null;
  return { data: m[1].trim(), valor: parseFloat(m[2].trim()), identificador: m[3].trim(), descricao: m[4].trim() };
}

function parseDataBr(data) {
  const m = data.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  return `${m[3]}-${m[2]}-${m[1]}`;
}

// Extrai nome do pagador e os 6 dígitos centrais do CPF mascarado (ex: "•••.488.328-••" -> "488328")
function extrairPagador(descricao) {
  const m = descricao.match(/^Transferência (?:recebida pelo Pix|Recebida) - (.+?) -/);
  if (!m) return null;
  const nome = m[1].trim();
  const cpfMatch = descricao.match(/•{2,3}\.(\d{3}\.\d{3})-•{2}/);
  const cpfParcial = cpfMatch ? cpfMatch[1].replace(/\./g, '') : null;
  return { nome, cpfParcial };
}

function lerTransacoes(caminho) {
  const linhas = fs.readFileSync(caminho, 'utf-8').split('\n').map(l => l.trim()).filter(Boolean);
  const transacoes = [];
  for (const linha of linhas.slice(1)) { // pula cabeçalho
    const row = parseCsvLine(linha);
    if (!row || Number.isNaN(row.valor)) continue;
    if (row.valor <= 0) continue; // só recebimentos (dinheiro entrando)
    const pagador = extrairPagador(row.descricao);
    if (!pagador) continue; // não é um Pix recebido identificável (ex: rendimento, estorno)
    transacoes.push({
      data: parseDataBr(row.data),
      valor: row.valor,
      nome: pagador.nome,
      cpfParcial: pagador.cpfParcial,
    });
  }
  return transacoes;
}

// ---------------------------------------------------------------------------
// 2. Busca pagamentos pendentes no Notion
// ---------------------------------------------------------------------------

async function buscarPendentes() {
  const pendentes = [];
  let cursor;
  do {
    const body = { filter: { property: 'Status', select: { equals: 'Pendente' } }, page_size: 100 };
    if (cursor) body.start_cursor = cursor;
    const res = await fetch(`https://api.notion.com/v1/data_sources/${PAGAMENTOS_DATA_SOURCE_ID}/query`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${NOTION_TOKEN}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2025-09-03',
      },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(`Erro ao buscar pendentes: ${JSON.stringify(data)}`);
    for (const page of data.results) {
      const p = page.properties;
      pendentes.push({
        id: page.id,
        nome: p['Nome']?.title?.[0]?.plain_text || '',
        professor: p['Professor']?.select?.name || null,
        turma: p['Turma']?.select?.name || null,
        mes: p['Mês']?.select?.name || null,
        aPagar: p['À Pagar']?.number ?? null,
      });
    }
    cursor = data.has_more ? data.next_cursor : null;
  } while (cursor);
  return pendentes;
}

// ---------------------------------------------------------------------------
// 3. Conciliação: nome (fuzzy) + valor exato
// ---------------------------------------------------------------------------

const MESES_ORDEM = ['Jan/26', 'Fev/26', 'Mar/26', 'Abr/26', 'Mai/26', 'Jun/26', 'Jul/26', 'Ago/26', 'Set/26', 'Out/26', 'Nov/26', 'Dez/26'];

// Remove anotações entre parênteses do nome (ex: "Caetano (Laís)" -> "Caetano", onde "Laís"
// é o nome da mãe — comum nas turmas infantis) antes de comparar nomes.
function stripParenteticoExtrato(nome) {
  return String(nome || '').replace(/\s*\([^)]*\)\s*/g, ' ').trim();
}

function candidatosPorNome(pendentes, nomeTransacao) {
  const key = normalizar(stripParenteticoExtrato(nomeTransacao));
  let candidatos = pendentes.filter(p => normalizar(stripParenteticoExtrato(p.nome)) === key);
  if (candidatos.length > 0) return candidatos;

  // fuzzy: primeiro nome + último sobrenome (cobre nome do meio adicionado/removido/abreviado)
  const tokens = key.split(' ').filter(Boolean);
  if (tokens.length >= 2) {
    const flKey = `${tokens[0]} ${tokens[tokens.length - 1]}`;
    candidatos = pendentes.filter(p => {
      const pTokens = normalizar(stripParenteticoExtrato(p.nome)).split(' ').filter(Boolean);
      if (pTokens.length < 2) return false;
      return `${pTokens[0]} ${pTokens[pTokens.length - 1]}` === flKey;
    });
  }
  return candidatos;
}

function conciliar(transacoes, pendentes) {
  const usados = new Set(); // ids de pendentes já usados nessa rodada, pra não casar 2x
  const resultados = [];

  for (const tx of transacoes) {
    const candidatosNome = candidatosPorNome(pendentes, tx.nome).filter(p => !usados.has(p.id));

    if (candidatosNome.length === 0) {
      resultados.push({ tx, status: 'sem_candidato' });
      continue;
    }

    const candidatosValor = candidatosNome.filter(p => p.aPagar === tx.valor);

    if (candidatosValor.length === 0) {
      resultados.push({ tx, status: 'valor_nao_bate', candidatos: candidatosNome });
      continue;
    }

    if (candidatosValor.length === 1) {
      usados.add(candidatosValor[0].id);
      resultados.push({ tx, status: 'ok', pendente: candidatosValor[0] });
      continue;
    }

    // múltiplos pendentes com mesmo nome e mesmo valor (ex: mensalidade igual em meses
    // seguidos) -> assume que é o mês em aberto mais antigo
    candidatosValor.sort((a, b) => MESES_ORDEM.indexOf(a.mes) - MESES_ORDEM.indexOf(b.mes));
    const escolhido = candidatosValor[0];
    usados.add(escolhido.id);
    resultados.push({ tx, status: 'ambiguo_resolvido', pendente: escolhido, outrosCandidatos: candidatosValor.slice(1) });
  }

  return resultados;
}

// ---------------------------------------------------------------------------
// 4. Atualização no Notion
// ---------------------------------------------------------------------------

async function marcarComoPago(pendenteId, dataPagamento, valorPago) {
  const res = await fetch(`https://api.notion.com/v1/pages/${pendenteId}`, {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${NOTION_TOKEN}`,
      'Content-Type': 'application/json',
      'Notion-Version': '2025-09-03',
    },
    body: JSON.stringify({
      properties: {
        'Status': { select: { name: 'Pago' } },
        'Data Pagamento': { date: { start: dataPagamento } },
        'Valor Pago': { number: valorPago },
      },
    }),
  });
  if (!res.ok) {
    const err = await res.text();
    console.error(`ERRO ao atualizar ${pendenteId}: ${err}`);
    return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// 5. Main
// ---------------------------------------------------------------------------

async function main() {
  console.log(DRY_RUN ? '=== MODO DRY-RUN (nada será atualizado no Notion) ===\n' : '=== MODO CONFIRM (vai atualizar registros de verdade!) ===\n');

  const transacoes = lerTransacoes(csvPath);
  console.log(`Extrato: ${transacoes.length} recebimentos identificados.`);

  const pendentes = await buscarPendentes();
  console.log(`Notion: ${pendentes.length} pagamentos pendentes no banco Pagamentos.\n`);

  const resultados = conciliar(transacoes, pendentes);

  const ok = resultados.filter(r => r.status === 'ok' || r.status === 'ambiguo_resolvido');
  const valorNaoBate = resultados.filter(r => r.status === 'valor_nao_bate');
  const semCandidato = resultados.filter(r => r.status === 'sem_candidato');

  console.log(`--- ${ok.length} pagamentos identificados com confiança ---`);
  for (const r of ok) {
    const tag = r.status === 'ambiguo_resolvido' ? '  [ambíguo: mesmo nome/valor em >1 mês, usado o mais antigo]' : '';
    console.log(`${r.tx.data} | R$ ${r.tx.valor.toFixed(2).padStart(8)} | ${r.tx.nome} -> ${r.pendente.nome} | ${r.pendente.professor}/${r.pendente.turma} | ${r.pendente.mes}${tag}`);
  }

  if (valorNaoBate.length > 0) {
    console.log(`\n--- ${valorNaoBate.length} pagador encontrado mas VALOR não bate com nenhum pendente dela (revisar manualmente) ---`);
    for (const r of valorNaoBate) {
      const opcoes = r.candidatos.map(c => `${c.mes}:R$${c.aPagar}`).join(', ');
      console.log(`${r.tx.data} | R$ ${r.tx.valor.toFixed(2).padStart(8)} | ${r.tx.nome} | pendentes dela: ${opcoes}`);
    }
  }

  if (semCandidato.length > 0) {
    console.log(`\n--- ${semCandidato.length} pagador SEM nenhum pendente correspondente (pode ser de outra modalidade/professor ainda não migrado, ou não relacionado a mensalidade) ---`);
    for (const r of semCandidato) {
      console.log(`${r.tx.data} | R$ ${r.tx.valor.toFixed(2).padStart(8)} | ${r.tx.nome}${r.tx.cpfParcial ? ` | CPF: •••.${r.tx.cpfParcial.slice(0,3)}.${r.tx.cpfParcial.slice(3)}-••` : ''}`);
    }
  }

  console.log(`\nResumo: ${ok.length} conciliados, ${valorNaoBate.length} valor não bate, ${semCandidato.length} sem candidato (de ${transacoes.length} recebimentos no extrato).`);

  if (DRY_RUN) {
    console.log('\nDry-run concluído. Rode com --confirm pra marcar os conciliados como Pago no Notion.');
    return;
  }

  console.log('\nAtualizando Notion...');
  let atualizados = 0;
  for (const r of ok) {
    const sucesso = await marcarComoPago(r.pendente.id, r.tx.data, r.tx.valor);
    if (sucesso) atualizados++;
    await new Promise(resolve => setTimeout(resolve, 350));
  }
  console.log(`\n=== Concluído: ${atualizados}/${ok.length} registros marcados como Pago ===`);
}

main().catch(e => {
  console.error('Erro fatal:', e);
  process.exit(1);
});
