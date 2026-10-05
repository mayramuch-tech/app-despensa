/**
 * Integração Bling API v3 (OAuth 2.0).
 * Credenciais e tokens ficam em Script Properties (nunca na planilha).
 */

const BLING_API = 'https://api.bling.com.br/Api/v3';
const BLING_AUTORIZAR = 'https://www.bling.com.br/Api/v3/oauth/authorize';
const BLING_TOKEN = 'https://api.bling.com.br/Api/v3/oauth/token';

function props_() { return PropertiesService.getScriptProperties(); }

/** Menu: grava client_id e client_secret do app criado na Central de Extensões do Bling. */
function configurarBling() {
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    ui.alert('Antes, publique o script como App da Web (Implantar ▸ Nova implantação ▸ App da Web). Veja o README, passo 3.');
    return;
  }
  const id = ui.prompt('Bling – Client ID', 'Cole o Client ID do seu aplicativo Bling:', ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK) return;
  const sec = ui.prompt('Bling – Client Secret', 'Cole o Client Secret:', ui.ButtonSet.OK_CANCEL);
  if (sec.getSelectedButton() !== ui.Button.OK) return;
  props_().setProperties({ BLING_CLIENT_ID: id.getResponseText().trim(), BLING_CLIENT_SECRET: sec.getResponseText().trim() });
  autorizarBling();
}

/** Menu: abre o link de autorização do Bling. */
function autorizarBling() {
  const p = props_();
  if (!p.getProperty('BLING_CLIENT_ID')) { configurarBling(); return; }
  const estado = Utilities.getUuid();
  p.setProperty('BLING_STATE', estado);
  const link = BLING_AUTORIZAR + '?response_type=code&client_id=' + encodeURIComponent(p.getProperty('BLING_CLIENT_ID')) +
    '&state=' + estado;
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px">' +
    '<p>URL de callback que deve estar cadastrada no app do Bling:</p>' +
    '<p style="word-break:break-all;background:#f1f5f9;padding:6px">' + ScriptApp.getService().getUrl() + '</p>' +
    '<p><a href="' + link + '" target="_blank" style="background:#ea580c;color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none">Autorizar no Bling</a></p>' +
    '<p>Depois de autorizar, volte aqui e rode <b>Cadência ▸ Testar conexão Bling</b>.</p></div>'
  ).setWidth(460).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Conectar Bling');
}

/** Callback OAuth do Bling (App da Web). */
function doGet(e) {
  const p = props_();
  const prm = (e && e.parameter) || {};
  if (!prm.code) return HtmlService.createHtmlOutput('Cadência integrada: aguardando autorização do Bling.');
  if (prm.state !== p.getProperty('BLING_STATE')) {
    return HtmlService.createHtmlOutput('❌ Estado inválido. Gere um novo link em Cadência ▸ Autorizar Bling.');
  }
  try {
    trocarToken_({ grant_type: 'authorization_code', code: prm.code });
    p.deleteProperty('BLING_STATE');
    return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif">✅ Bling conectado! Pode fechar esta aba.</h2>');
  } catch (err) {
    return HtmlService.createHtmlOutput('❌ Erro ao conectar: ' + escaparHtml_(err.message));
  }
}

function trocarToken_(payload) {
  const p = props_();
  const basic = Utilities.base64Encode(p.getProperty('BLING_CLIENT_ID') + ':' + p.getProperty('BLING_CLIENT_SECRET'));
  const resp = UrlFetchApp.fetch(BLING_TOKEN, {
    method: 'post',
    headers: { Authorization: 'Basic ' + basic, Accept: '1.0' },
    payload: payload,
    muteHttpExceptions: true,
  });
  const corpo = JSON.parse(resp.getContentText() || '{}');
  if (resp.getResponseCode() !== 200 || !corpo.access_token) {
    throw new Error('Bling OAuth ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 300));
  }
  p.setProperties({
    BLING_ACCESS_TOKEN: corpo.access_token,
    BLING_REFRESH_TOKEN: corpo.refresh_token,
    BLING_EXPIRA: String(Date.now() + (Number(corpo.expires_in || 21600) - 300) * 1000),
  });
  return corpo.access_token;
}

function tokenBling_() {
  const p = props_();
  if (!p.getProperty('BLING_REFRESH_TOKEN')) throw new Error('Bling não autorizado. Use Cadência ▸ Configurar Bling.');
  if (Date.now() < Number(p.getProperty('BLING_EXPIRA') || 0)) return p.getProperty('BLING_ACCESS_TOKEN');
  return trocarToken_({ grant_type: 'refresh_token', refresh_token: p.getProperty('BLING_REFRESH_TOKEN') });
}

/** GET na API v3 com renovação de token e respeito ao limite de 3 req/s. */
function blingGet_(caminho, params) {
  const qs = Object.keys(params || {}).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
  const url = BLING_API + caminho + (qs ? '?' + qs : '');
  for (let tentativa = 0; tentativa < 3; tentativa++) {
    Utilities.sleep(350);
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + tokenBling_(), Accept: 'application/json' },
      muteHttpExceptions: true,
    });
    const cod = resp.getResponseCode();
    if (cod === 200) return JSON.parse(resp.getContentText());
    if (cod === 401) { props_().setProperty('BLING_EXPIRA', '0'); continue; }
    if (cod === 429) { Utilities.sleep(1500 * (tentativa + 1)); continue; }
    if (cod === 404) return { data: [] };
    throw new Error('Bling ' + cod + ' em ' + caminho + ': ' + resp.getContentText().substring(0, 300));
  }
  throw new Error('Bling: falha após 3 tentativas em ' + caminho);
}

/** Lista paginada (100 por página). */
function blingListar_(caminho, params) {
  const todos = [];
  for (let pagina = 1; pagina <= 50; pagina++) {
    const r = blingGet_(caminho, Object.assign({}, params, { pagina: pagina, limite: 100 }));
    const dados = r.data || [];
    todos.push.apply(todos, dados);
    if (dados.length < 100) break;
  }
  return todos;
}

/** Situação pode vir como objeto {id, valor} (pedidos) ou como texto/número (propostas). */
function situacaoBling_(s) {
  if (s === null || s === undefined) return '';
  if (typeof s === 'object') return String(s.id !== undefined ? s.id : (s.valor || s.nome || ''));
  return String(s);
}

function contatoBling_(id) {
  if (!id) return {};
  const r = blingGet_('/contatos/' + id);
  return r.data || {};
}

/** Menu: confirma conexão e mostra a empresa conectada. */
function testarBling() {
  try {
    const r = blingGet_('/empresas/me/dados-basicos');
    SpreadsheetApp.getUi().alert('✅ Bling conectado: ' + ((r.data && (r.data.nome || r.data.razaoSocial)) || 'OK'));
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

/** Menu: grava na aba Situacoes_Bling todos os módulos e situações (para preencher a Config). */
function listarSituacoesBling() {
  const linhas = [];
  (blingGet_('/situacoes/modulos').data || []).forEach(m => {
    (blingGet_('/situacoes/modulos/' + m.id).data || []).forEach(s => {
      linhas.push([m.id, m.nome || m.descricao, s.id, s.nome]);
    });
  });
  const sh = aba_(ABA.SITUACOES);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (linhas.length) sh.getRange(2, 1, linhas.length, 4).setValues(linhas);
  SpreadsheetApp.getActive().toast(linhas.length + ' situações listadas.', 'Bling', 5);
}

/**
 * Sincroniza pedidos de venda e propostas comerciais:
 *  • novos → inscreve na cadência correspondente;
 *  • já inscritos → atualiza situação e encerra (cancelado / aprovado / não aprovado).
 */
function sincronizarBling_(tab) {
  const c = cfg_();
  const hoje = new Date();
  const ini = new Date(); ini.setDate(ini.getDate() - Number(c.BLING_JANELA_DIAS || 15));
  const periodo = { dataInicial: dataISO_(ini), dataFinal: dataISO_(hoje) };

  const porRef = {};
  tab.linhas.forEach(l => { if (l.Ref_Bling) porRef[String(l.Ref_Bling)] = l; });
  const novos = [];

  // ── Pedidos de venda ──
  const entrada = cfgLista_('BLING_PEDIDO_SITUACOES_ENTRADA');
  const paradaPed = cfgLista_('BLING_PEDIDO_SITUACOES_PARADA');
  blingListar_('/pedidos/vendas', periodo).forEach(p => {
    const ref = 'PED-' + p.id, sit = situacaoBling_(p.situacao);
    const ex = porRef[ref];
    if (ex) {
      if (String(ex.Situacao_Bling) !== sit) { ex.Situacao_Bling = sit; ex._sujo = true; }
      if (paradaPed.indexOf(normaliza_(sit)) >= 0 && STATUS_FINAIS.indexOf(ex.Status) < 0) {
        encerrar_(ex, STATUS.CANCELADO, 'Pedido ' + (p.numero || p.id) + ' na situação de parada ' + sit);
      }
      return;
    }
    if (entrada.indexOf(normaliza_(sit)) < 0) return;
    const ct = contatoBling_(p.contato && p.contato.id);
    novos.push(montarInscricao_(c.CADENCIA_PEDIDO, 'BLING_PEDIDO', ref, sit, ct, p.contato, {
      numero_pedido: p.numero || p.id,
      total: formatarMoeda_(p.total),
      data_pedido: p.data || '',
      id_pedido_bling: p.id,
    }));
  });

  // ── Propostas comerciais ──
  const ganho = cfgLista_('BLING_PROPOSTA_SITUACOES_GANHO');
  const perda = cfgLista_('BLING_PROPOSTA_SITUACOES_PERDA');
  const ignorar = cfgLista_('BLING_PROPOSTA_SITUACOES_IGNORAR');
  blingListar_('/propostas-comerciais', periodo).forEach(pr => {
    const ref = 'PROP-' + pr.id, sit = situacaoBling_(pr.situacao), n = normaliza_(sit);
    const ex = porRef[ref];
    if (ex) {
      if (String(ex.Situacao_Bling) !== sit) { ex.Situacao_Bling = sit; ex._sujo = true; }
      if (STATUS_FINAIS.indexOf(ex.Status) < 0) {
        if (ganho.indexOf(n) >= 0) encerrar_(ex, STATUS.GANHO, 'Proposta ' + (pr.numero || pr.id) + ': ' + sit);
        else if (perda.indexOf(n) >= 0) encerrar_(ex, STATUS.PERDIDO, 'Proposta ' + (pr.numero || pr.id) + ': ' + sit);
      }
      return;
    }
    if (ganho.concat(perda, ignorar).indexOf(n) >= 0) return;
    const ct = contatoBling_(pr.contato && pr.contato.id);
    novos.push(montarInscricao_(c.CADENCIA_PROPOSTA, 'BLING_PROPOSTA', ref, sit, ct, pr.contato, {
      numero_proposta: pr.numero || pr.id,
      total: formatarMoeda_(pr.total || pr.totalProdutos),
      data_proposta: pr.data || '',
    }));
  });

  if (novos.length) {
    anexarLinhas_(ABA.CONTATOS, novos.map(o => CAB.CONTATOS.map(k => o[k] === undefined ? '' : o[k])));
    novos.forEach(o => log_(o.ID, 'INSCRITO', o.Cadencia + ' via ' + o.Origem + ' (' + o.Ref_Bling + ')'));
  }
  return novos.length;
}

function montarInscricao_(cadencia, origem, ref, sit, ct, contatoResumo, vars) {
  const nome = ct.nome || (contatoResumo && contatoResumo.nome) || '';
  return {
    ID: novoId_(),
    Cadencia: cadencia,
    Status: '',               // vazio → o motor inicializa no próximo ciclo
    Passo: '',
    Nome: nome,
    Empresa: ct.fantasia || (ct.tipo === 'J' ? nome : ''),
    Email: String(ct.email || '').split(/[;,]/)[0].trim(),
    Telefone: ct.celular || ct.telefone || '',
    Origem: origem,
    Ref_Bling: ref,
    Situacao_Bling: sit,
    Vars_JSON: JSON.stringify(vars),
  };
}

/** Busca o código de rastreio no pedido (usado só quando o modelo pede {{rastreio}}). */
function rastreioPedido_(idPedido) {
  try {
    const p = blingGet_('/pedidos/vendas/' + idPedido).data || {};
    const vols = (p.transporte && p.transporte.volumes) || [];
    return vols.map(v => v.codigoRastreamento).filter(String).join(', ');
  } catch (e) {
    return '';
  }
}
