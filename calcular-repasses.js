// Calcula o repasse (pagamento) devido a cada professor por mês e grava no banco "Repasses" do Notion.
//
// Regras:
// - Titzi, Gustra, Guilherme: R$80 por aula dada. Calculado pelo calendário (dias da semana da(s)
//   turma(s) no mês, menos feriados nacionais/estaduais/municipais de SP) — o app de Presença é
//   recente demais pra ter histórico suficiente.
// - Gabi, Talita: 65% do que suas alunas efetivamente pagaram (soma "Valor Pago" dos registros com
//   Status=Pago no banco Recebimentos).
//
// Opcionalmente aceita um extrato bancário (CSV Nubank) como argumento pra comparar o calculado
// contra o que já foi pago de verdade (transações de saída/Pix enviado), como conferência.
//
// Uso:
//   railway run node calcular-repasses.js [extrato.csv] --dry-run
//   railway run node calcular-repasses.js [extrato.csv] --confirm

const fs = require('fs');

const NOTION_TOKEN = process.env.NOTION_TOKEN;
const RECEBIMENTOS_DB = '5b85a30acc9f406d81c082bdf521d5a8';
const REPASSES_DB = '4cf21b7e3b824eec9a31420ce02be5a1';

const args = process.argv.slice(2);
const csvPath = args.find(a => !a.startsWith('--'));
const DRY_RUN = !args.includes('--confirm');

const VALOR_POR_AULA = 80;
const PERCENTUAL = 0.65;

// Turmas reais por professor (mesma fonte da migração de Recebimentos)
const CALENDARIO_POR_AULA = {
  Gustra: [{ dia: 'quarta', turmas: 2 }],       // Quarta 18h + Quarta 19h
  Guilherme: [{ dia: 'quinta', turmas: 1 }],    // Quinta 8h
  Titzi: [{ dia: 'terca', turmas: 1 }, { dia: 'quarta', turmas: 1 }], // Terca 18h + Quarta 9h30
};
const PROFESSORES_POR_PERCENTUAL = { Gabi: PERCENTUAL, Talita: PERCENTUAL };

// Nomes como aparecem nos Pix de saída do extrato (confirmados com o Fábio)
const NOME_EXTRATO_POR_PROFESSOR = {
  Gustra: 'gustavo henrique das dores',
  Guilherme: 'guilherme fillippi guerra',
  Titzi: 'titziane marques',
  Talita: 'talita silva',
  Gabi: 'gabriela gomes de sousa',
};

const MESES_LABEL = { 1: 'Jan/26', 2: 'Fev/26', 3: 'Mar/26', 4: 'Abr/26', 5: 'Mai/26', 6: 'Jun/26', 7: 'Jul/26', 8: 'Ago/26', 9: 'Set/26', 10: 'Out/26', 11: 'Nov/26', 12: 'Dez/26' };
const MESES_ORDEM = Object.values(MESES_LABEL);
const MES_ATUAL_LIMITE = 8; // não calcula meses futuros a mais do que já migramos (Jan-Ago/26)

function normalizar(str) {
  return String(str || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim().replace(/\s+/g, ' ');
}

// ---------------------------------------------------------------------------
// 1. Feriados (nacionais + SP estadual + SP municipal), mesma lógica do server.js
// ---------------------------------------------------------------------------

function calcularPascoa(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(ano, mes - 1, dia));
}

function formatarDataStrUTC(data) {
  return data.toISOString().split('T')[0];
}

function getFeriadosDoAno(ano) {
  const feriados = new Set();
  feriados.add(ano + '-01-01');
  feriados.add(ano + '-04-21');
  feriados.add(ano + '-05-01');
  feriados.add(ano + '-09-07');
  feriados.add(ano + '-10-12');
  feriados.add(ano + '-11-02');
  feriados.add(ano + '-11-15');
  feriados.add(ano + '-11-20');
  feriados.add(ano + '-12-25');
  feriados.add(ano + '-07-09');
  feriados.add(ano + '-01-25');
  const pascoa = calcularPascoa(ano);
  const addDias = (data, dias) => new Date(data.getTime() + dias * 24 * 60 * 60000);
  feriados.add(formatarDataStrUTC(addDias(pascoa, -48)));
  feriados.add(formatarDataStrUTC(addDias(pascoa, -47)));
  feriados.add(formatarDataStrUTC(addDias(pascoa, -2)));
  feriados.add(formatarDataStrUTC(addDias(pascoa, 60)));
  return feriados;
}

// ---------------------------------------------------------------------------
// 2. Cálculo por aula (calendário menos feriados)
// ---------------------------------------------------------------------------

const DIA_INDEX = { segunda: 1, terca: 2, quarta: 3, quinta: 4, sexta: 5, sabado: 6, domingo: 0 };

function contarDiaDaSemanaNoMes(ano, mesNum, diaSemanaNome, feriados) {
  const diaAlvo = DIA_INDEX[diaSemanaNome];
  const datas = [];
  const ultimoDia = new Date(Date.UTC(ano, mesNum, 0)).getUTCDate();
  for (let d = 1; d <= ultimoDia; d++) {
    const data = new Date(Date.UTC(ano, mesNum - 1, d));
    if (data.getUTCDay() === diaAlvo) {
      const dataStr = data.toISOString().slice(0, 10);
      if (!feriados.has(dataStr)) datas.push(dataStr);
    }
  }
  return datas;
}

function calcularPorAula(ano) {
  const feriados = getFeriadosDoAno(ano);
  const resultado = [];
  for (const [professor, blocos] of Object.entries(CALENDARIO_POR_AULA)) {
    for (let mesNum = 1; mesNum <= MES_ATUAL_LIMITE; mesNum++) {
      let qtdAulas = 0;
      const detalhe = [];
      for (const bloco of blocos) {
        const datas = contarDiaDaSemanaNoMes(ano, mesNum, bloco.dia, feriados);
        qtdAulas += datas.length * bloco.turmas;
        detalhe.push({ dia: bloco.dia, turmas: bloco.turmas, datas });
      }
      resultado.push({
        professor,
        mes: MESES_LABEL[mesNum],
        metodo: 'por_aula',
        qtdAulas,
        valorRepasse: qtdAulas * VALOR_POR_AULA,
        detalhe,
        recebimentosIds: [],
      });
    }
  }
  return resultado;
}

// ---------------------------------------------------------------------------
// 3. Cálculo por percentual (Notion Recebimentos)
// ---------------------------------------------------------------------------

async function notionQuery(databaseId, filtro) {
  const registros = [];
  let cursor;
  do {
    const body = { page_size: 100 };
    if (filtro) body.filter = filtro;
    if (cursor) body.start_cursor = cursor;
    const res = await fetch('https://api.notion.com/v1/databases/' + databaseId + '/query', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const d = await res.json();
    if (!res.ok) throw new Error('Erro ao consultar ' + databaseId + ': ' + JSON.stringify(d));
    registros.push(...d.results);
    cursor = d.has_more ? d.next_cursor : null;
  } while (cursor);
  return registros;
}

async function calcularPorPercentual() {
  const professores = Object.keys(PROFESSORES_POR_PERCENTUAL);
  const filtro = {
    and: [
      { or: professores.map(p => ({ property: 'Professor', select: { equals: p } })) },
      { property: 'Status', select: { equals: 'Pago' } },
    ],
  };
  const paginas = await notionQuery(RECEBIMENTOS_DB, filtro);

  const grupos = {};
  for (const pagina of paginas) {
    const p = pagina.properties;
    const professor = p['Professor']?.select?.name;
    const mes = p['Mês']?.select?.name;
    const valorPago = p['Valor Pago']?.number || 0;
    if (!professor || !mes) continue;
    const chave = `${professor}|${mes}`;
    if (!grupos[chave]) grupos[chave] = { total: 0, ids: [] };
    grupos[chave].total += valorPago;
    grupos[chave].ids.push(pagina.id);
  }

  const resultado = [];
  for (const [chave, g] of Object.entries(grupos)) {
    const [professor, mes] = chave.split('|');
    resultado.push({
      professor,
      mes,
      metodo: 'percentual',
      totalRecebido: Math.round(g.total * 100) / 100,
      valorRepasse: Math.round(g.total * PROFESSORES_POR_PERCENTUAL[professor] * 100) / 100,
      recebimentosIds: g.ids,
    });
  }
  return resultado;
}

// ---------------------------------------------------------------------------
// 4. Extrato (saídas) — só pra conferência, opcional
// ---------------------------------------------------------------------------

function identificarProfessorPix(nomeExtrato) {
  const key = normalizar(nomeExtrato);
  for (const [professor, padrao] of Object.entries(NOME_EXTRATO_POR_PROFESSOR)) {
    if (key.startsWith(padrao)) return professor;
  }
  return null;
}

function lerSaidasExtrato(caminho) {
  if (!caminho) return [];
  const linhas = fs.readFileSync(caminho, 'utf-8').split('\n').map(l => l.trim()).filter(Boolean);
  const saidas = [];
  for (const linha of linhas.slice(1)) {
    const m = linha.match(/^([^,]*),([^,]*),([^,]*),(.*)$/);
    if (!m) continue;
    const valor = parseFloat(m[2].trim());
    if (Number.isNaN(valor) || valor >= 0) continue; // só saídas
    const descricao = m[4].trim();
    const nomeMatch = descricao.match(/^Transferência enviada pelo Pix - (.+?) -/);
    if (!nomeMatch) continue;
    const dataMatch = m[1].trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!dataMatch) continue;
    const professor = identificarProfessorPix(nomeMatch[1]);
    if (!professor) continue;
    const mesNum = parseInt(dataMatch[2], 10);
    saidas.push({ professor, mes: MESES_LABEL[mesNum], valor: Math.abs(valor), data: `${dataMatch[3]}-${dataMatch[2]}-${dataMatch[1]}` });
  }

  const porProfessorMes = {};
  for (const s of saidas) {
    const chave = `${s.professor}|${s.mes}`;
    porProfessorMes[chave] = (porProfessorMes[chave] || 0) + s.valor;
  }
  return porProfessorMes;
}

// ---------------------------------------------------------------------------
// 5. Repasses existentes no Notion (pra atualizar em vez de duplicar)
// ---------------------------------------------------------------------------

async function buscarRepassesExistentes() {
  const paginas = await notionQuery(REPASSES_DB, null);
  const mapa = {};
  for (const pagina of paginas) {
    const p = pagina.properties;
    const professor = p['Professor']?.select?.name;
    const mes = p['Mês']?.select?.name;
    if (!professor || !mes) continue;
    mapa[`${professor}|${mes}`] = {
      id: pagina.id,
      valorRepasseAtual: p['Valor Repasse']?.number ?? null,
      status: p['Status']?.select?.name || null,
    };
  }
  return mapa;
}

async function criarOuAtualizarRepasse(calc, existente) {
  const nome = `${calc.professor} — ${calc.mes}`;
  const properties = {
    'Nome': { title: [{ text: { content: nome } }] },
    'Professor': { select: { name: calc.professor } },
    'Mês': { select: { name: calc.mes } },
    'Valor Repasse': { number: calc.valorRepasse },
  };
  if (calc.recebimentosIds.length > 0) {
    properties['Pagamentos do Mês'] = { relation: calc.recebimentosIds.map(id => ({ id })) };
  }

  if (existente) {
    const res = await fetch('https://api.notion.com/v1/pages/' + existente.id, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties }),
    });
    if (!res.ok) throw new Error('Erro ao atualizar ' + nome + ': ' + JSON.stringify(await res.json()));
  } else {
    properties['Status'] = { select: { name: 'Pendente' } };
    const res = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + NOTION_TOKEN, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: { database_id: REPASSES_DB }, properties }),
    });
    if (!res.ok) throw new Error('Erro ao criar ' + nome + ': ' + JSON.stringify(await res.json()));
  }
}

// ---------------------------------------------------------------------------
// 6. Main
// ---------------------------------------------------------------------------

async function main() {
  console.log(DRY_RUN ? '=== MODO DRY-RUN (nada será gravado no Notion) ===\n' : '=== MODO CONFIRM (vai gravar no Notion!) ===\n');

  const porAula = calcularPorAula(2026);
  const porPercentual = await calcularPorPercentual();
  const existentes = await buscarRepassesExistentes();
  const saidasExtrato = lerSaidasExtrato(csvPath);

  if (csvPath) console.log(`Extrato: ${csvPath}\n`);

  const todos = [...porAula, ...porPercentual].sort((a, b) => {
    const profCmp = a.professor.localeCompare(b.professor, 'pt-BR');
    if (profCmp !== 0) return profCmp;
    return MESES_ORDEM.indexOf(a.mes) - MESES_ORDEM.indexOf(b.mes);
  });

  console.log('--- Cálculo por professor/mês ---');
  for (const calc of todos) {
    const existente = existentes[`${calc.professor}|${calc.mes}`];
    const extratoValor = saidasExtrato[`${calc.professor}|${calc.mes}`];
    const linhaBase = calc.metodo === 'por_aula'
      ? `${calc.qtdAulas} aulas × R$${VALOR_POR_AULA}`
      : `R$${calc.totalRecebido.toFixed(2)} recebido × 65%`;
    let linha = `${calc.professor.padEnd(10)} ${calc.mes} | ${linhaBase} = R$ ${calc.valorRepasse.toFixed(2)}`;
    if (extratoValor != null) {
      const diffExtrato = (calc.valorRepasse - extratoValor).toFixed(2);
      linha += ` | extrato pagou: R$${extratoValor.toFixed(2)}${Math.abs(diffExtrato) >= 0.01 ? ` ⚠ DIF R$${diffExtrato}` : ' ✅'}`;
    }
    if (existente) {
      linha += ` | Notion atual: R$${existente.valorRepasseAtual} (${existente.status})`;
    } else {
      linha += ' | NOVO';
    }
    console.log(linha);
  }

  console.log(`\nTotal: ${todos.length} registros (${porAula.length} por aula, ${porPercentual.length} por percentual)`);

  if (DRY_RUN) {
    console.log('\nDry-run concluído. Rode com --confirm pra gravar/atualizar no Notion.');
    return;
  }

  console.log('\nGravando no Notion...');
  let ok = 0;
  for (const calc of todos) {
    const existente = existentes[`${calc.professor}|${calc.mes}`];
    try {
      await criarOuAtualizarRepasse(calc, existente);
      ok++;
    } catch (e) {
      console.error('ERRO:', e.message);
    }
    await new Promise(r => setTimeout(r, 350));
  }
  console.log(`\n=== Concluído: ${ok}/${todos.length} registros gravados ===`);
}

main().catch(e => {
  console.error('Erro fatal:', e);
  process.exit(1);
});
