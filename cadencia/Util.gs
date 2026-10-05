/** Utilitários: planilha, datas, telefones, modelos e log. */

function aba_(nome) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada. Rode Cadência ▸ Instalar planilha.');
  return sh;
}

/**
 * Lê uma aba como lista de objetos { campo: valor, _linha: n }.
 * Os nomes dos campos vêm da linha 1.
 */
function lerTabela_(nome) {
  const sh = aba_(nome);
  const ultLinha = sh.getLastRow(), ultCol = sh.getLastColumn();
  if (ultLinha < 1) return { sh: sh, cab: [], linhas: [] };
  const valores = sh.getRange(1, 1, ultLinha, ultCol).getValues();
  const cab = valores[0].map(String);
  const linhas = valores.slice(1).map((r, i) => {
    const o = { _linha: i + 2 };
    cab.forEach((c, j) => { o[c] = r[j]; });
    return o;
  });
  return { sh: sh, cab: cab, linhas: linhas };
}

/** Grava de volta as linhas alteradas (marcadas com _sujo). */
function gravarTabela_(tab) {
  tab.linhas.filter(l => l._sujo).forEach(l => {
    tab.sh.getRange(l._linha, 1, 1, tab.cab.length).setValues([tab.cab.map(c => l[c] === undefined ? '' : l[c])]);
    l._sujo = false;
  });
}

function anexarLinhas_(nome, linhas) {
  if (!linhas.length) return;
  const sh = aba_(nome);
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
}

function log_(contatoId, tipo, detalhe) {
  try {
    anexarLinhas_(ABA.LOG, [[new Date(), contatoId || '', tipo, String(detalhe || '').substring(0, 2000)]]);
  } catch (e) {
    console.error('Falha ao registrar log: ' + e);
  }
}

/** Texto sem acento, minúsculo e sem espaços nas pontas — para comparar situações. */
function normaliza_(s) {
  return String(s === null || s === undefined ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
}

/** Telefone BR → só dígitos com DDI 55. Retorna '' se inválido. */
function telefoneWpp_(tel) {
  let d = String(tel || '').replace(/\D/g, '');
  if (!d) return '';
  d = d.replace(/^0+/, '');
  if (d.length === 10 || d.length === 11) d = '55' + d;
  return d.length >= 12 && d.length <= 13 ? d : '';
}

function linkWhatsApp_(tel, mensagem) {
  const n = telefoneWpp_(tel);
  return n ? 'https://wa.me/' + n + '?text=' + encodeURIComponent(mensagem) : '';
}

function ehDiaUtil_(d) {
  const dia = d.getDay();
  return dia !== 0 && dia !== 6;
}

/** Soma dias corridos ou úteis, conforme Config SO_DIAS_UTEIS. */
function somarDias_(base, dias) {
  const d = new Date(base.getTime());
  if (!cfgSim_('SO_DIAS_UTEIS')) {
    d.setDate(d.getDate() + Number(dias || 0));
    return d;
  }
  let restante = Number(dias || 0);
  while (restante > 0) {
    d.setDate(d.getDate() + 1);
    if (ehDiaUtil_(d)) restante--;
  }
  while (!ehDiaUtil_(d)) d.setDate(d.getDate() + 1);
  return d;
}

/** Está dentro da janela de envio (horário comercial e, se configurado, dia útil)? */
function dentroDoHorario_(agora) {
  const c = cfg_();
  const h = agora.getHours();
  if (h < Number(c.HORA_INICIO) || h >= Number(c.HORA_FIM)) return false;
  if (cfgSim_('SO_DIAS_UTEIS') && !ehDiaUtil_(agora)) return false;
  return true;
}

/** Substitui {{variavel}} pelos valores do contato + variáveis globais. */
function renderizar_(modelo, vars) {
  return String(modelo || '').replace(/\{\{\s*([\w]+)\s*\}\}/g, (m, chave) => {
    const v = vars[chave.toLowerCase()];
    return v === undefined || v === null ? '' : String(v);
  });
}

function escaparHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Texto simples → HTML com quebras de linha e links clicáveis. */
function textoParaHtml_(texto) {
  return escaparHtml_(texto)
    .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>')
    .split(/\n{2,}/).map(p => '<p>' + p.replace(/\n/g, '<br>') + '</p>').join('');
}

function primeiroNome_(nome) {
  const p = String(nome || '').trim().split(/\s+/)[0] || '';
  return p ? p.charAt(0).toUpperCase() + p.slice(1).toLowerCase() : '';
}

function formatarMoeda_(v) {
  const n = Number(v);
  return isNaN(n) ? String(v || '') : n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function dataISO_(d) {
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function novoId_() {
  return 'C' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyMMddHHmmss') +
    Math.floor(Math.random() * 1000).toString().padStart(3, '0');
}
