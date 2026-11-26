/**
 * Integração Wix eCommerce — recuperação de carrinhos abandonados.
 * API: POST https://www.wixapis.com/ecom/v1/abandoned-checkout/query
 * Autenticação: API key da conta Wix (header Authorization) + wix-site-id.
 *
 *  • Carrinho ABANDONADO recente  → cadência CARRINHO_ABANDONADO
 *  • Carrinho ABANDONADO antigo   → cadência CARRINHO_RETOMADA (abordagem de retomada)
 *  • Carrinho RECUPERADO          → encerra como GANHO
 */

const WIX_API = 'https://www.wixapis.com';

/** Menu: grava API key e ID do site (Script Properties, nunca na planilha). */
function configurarWix() {
  const ui = SpreadsheetApp.getUi();
  const k = ui.prompt('Wix – API key',
    'Gere em manage.wix.com/account/api-keys com a permissão "Wix Stores / eCommerce – Read Orders". Cole a chave:',
    ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK) return;
  const s = ui.prompt('Wix – ID do site',
    'Cole o ID do site (está na URL do painel: manage.wix.com/dashboard/<ID-DO-SITE>/...):', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  props_().setProperties({ WIX_API_KEY: k.getResponseText().trim(), WIX_SITE_ID: s.getResponseText().trim() });
  testarWix();
}

function wixConfigurado_() {
  return !!(props_().getProperty('WIX_API_KEY') && props_().getProperty('WIX_SITE_ID'));
}

function wixPost_(caminho, corpo) {
  const p = props_();
  for (let tentativa = 0; tentativa < 3; tentativa++) {
    const resp = UrlFetchApp.fetch(WIX_API + caminho, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: p.getProperty('WIX_API_KEY'), 'wix-site-id': p.getProperty('WIX_SITE_ID') },
      payload: JSON.stringify(corpo),
      muteHttpExceptions: true,
    });
    const cod = resp.getResponseCode();
    if (cod === 200) return JSON.parse(resp.getContentText());
    if (cod === 429 || cod >= 500) { Utilities.sleep(1500 * (tentativa + 1)); continue; }
    throw new Error('Wix ' + cod + ' em ' + caminho + ': ' + resp.getContentText().substring(0, 300));
  }
  throw new Error('Wix: falha após 3 tentativas em ' + caminho);
}

/** Lista carrinhos (abandonados e recuperados) criados a partir de `desde`, com paginação por cursor. */
function wixCarrinhos_(desde) {
  const todos = [];
  let query = {
    filter: { createdDate: { $gte: desde.toISOString() }, status: { $in: ['ABANDONED', 'RECOVERED'] } },
    sort: [{ fieldName: 'createdDate', order: 'ASC' }],
    cursorPaging: { limit: 100 },
  };
  for (let pagina = 0; pagina < 50; pagina++) {
    const r = wixPost_('/ecom/v1/abandoned-checkout/query', { query: query });
    todos.push.apply(todos, r.abandonedCheckouts || []);
    const prox = r.metadata && r.metadata.cursors && r.metadata.cursors.next;
    if (!prox) break;
    query = { cursorPaging: { limit: 100, cursor: prox } };
  }
  return todos;
}

/** Menu: confirma acesso e mostra quantos carrinhos abertos existem na janela configurada. */
function testarWix() {
  try {
    const desde = new Date(Date.now() - Number(cfg_().WIX_JANELA_DIAS || 60) * 86400000);
    const lista = wixCarrinhos_(desde);
    const abertos = lista.filter(c => c.status === 'ABANDONED').length;
    SpreadsheetApp.getUi().alert('✅ Wix conectado.\n\nÚltimos ' + cfg_().WIX_JANELA_DIAS + ' dias: ' +
      abertos + ' carrinho(s) em aberto e ' + (lista.length - abertos) + ' recuperado(s).');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

/** Menu: importa carrinhos agora, sem esperar o ciclo horário. */
function importarCarrinhosWix() {
  _configCache = null;
  const tab = lerTabela_(ABA.CONTATOS);
  const n = sincronizarWix_(tab);
  gravarTabela_(tab);
  SpreadsheetApp.getUi().alert(n + ' carrinho(s) novo(s) colocado(s) em cadência.\n' +
    'Rode "Processar agora" para gerar as mensagens.');
}

/** Texto curto com os itens: "Galão 5L Oxy Areia (1x), Plano Preventivo… (2x)". */
function produtosCarrinho_(c) {
  return (c.lineItems || []).map(i => {
    const nome = (i.productName && (i.productName.translated || i.productName.original)) || 'item';
    return nome + (Number(i.quantity) > 1 ? ' (' + i.quantity + 'x)' : '');
  }).join(', ');
}

const PROVEDORES_PESSOAIS = ['gmail.com', 'hotmail.com', 'outlook.com', 'live.com', 'yahoo.com', 'yahoo.com.br',
  'uol.com.br', 'bol.com.br', 'terra.com.br', 'icloud.com', 'ig.com.br', 'globo.com', 'msn.com'];

/** E-mail de domínio próprio (ex.: compras@empresa.com.br) → provável compra B2B. */
function ehEmailCorporativo_(email) {
  const dom = String(email || '').split('@')[1];
  return !!dom && PROVEDORES_PESSOAIS.indexOf(dom.toLowerCase()) < 0;
}

/**
 * Sincroniza carrinhos do Wix com a aba Contatos.
 * Um comprador com cadência de carrinho já ativa não recebe uma segunda (evita mensagens duplicadas).
 */
function sincronizarWix_(tab) {
  const c = cfg_();
  const agora = Date.now();
  const desde = new Date(agora - Number(c.WIX_JANELA_DIAS || 60) * 86400000);
  const esperaMs = Number(c.WIX_ESPERA_MIN || 60) * 60000;
  const limiteAntigoMs = Number(c.WIX_DIAS_CARRINHO_ANTIGO || 7) * 86400000;
  const cadencias = [c.CADENCIA_CARRINHO, c.CADENCIA_CARRINHO_ANTIGO];

  const porRef = {}, emCadencia = {};
  tab.linhas.forEach(l => {
    if (l.Ref_Bling) porRef[String(l.Ref_Bling)] = l;
    if (l.Email && cadencias.indexOf(l.Cadencia) >= 0 && STATUS_FINAIS.indexOf(l.Status) < 0) {
      emCadencia[String(l.Email).toLowerCase()] = true;
    }
  });

  const novos = [];
  wixCarrinhos_(desde).forEach(cart => {
    const ref = 'WIX-' + cart.id;
    const ex = porRef[ref];
    if (ex) {
      if (cart.status === 'RECOVERED' && STATUS_FINAIS.indexOf(ex.Status) < 0) {
        encerrar_(ex, STATUS.GANHO, 'Carrinho recuperado no Wix em ' + (cart.checkoutRecoveredDate || 'data n/d'));
      }
      if (String(ex.Situacao_Bling) !== cart.status) { ex.Situacao_Bling = cart.status; ex._sujo = true; }
      return;
    }
    if (cart.status !== 'ABANDONED') return;

    const criado = new Date(cart.createdDate).getTime();
    if (agora - criado < esperaMs) return; // dá tempo do cliente concluir sozinho

    const ct = cart.contactDetails || {};
    const email = String((cart.buyerInfo && cart.buyerInfo.email) || '').trim();
    if (!email && !ct.phone) return; // sem canal de contato
    if (email && emCadencia[email.toLowerCase()]) return;

    const antigo = agora - criado > limiteAntigoMs;
    const nome = [ct.firstName, ct.lastName].filter(String).join(' ').trim() || email.split('@')[0];
    const corporativo = ehEmailCorporativo_(email) || !!ct.company;
    novos.push({
      ID: novoId_(),
      Cadencia: antigo ? c.CADENCIA_CARRINHO_ANTIGO : c.CADENCIA_CARRINHO,
      Status: '',
      Nome: nome,
      Empresa: ct.company || (ehEmailCorporativo_(email) ? email.split('@')[1] : ''),
      Email: email,
      Telefone: ct.phone || '',
      Origem: 'WIX_CARRINHO',
      Ref_Bling: ref,
      Situacao_Bling: cart.status,
      Observacoes: antigo ? 'Carrinho antigo (' + Math.round((agora - criado) / 86400000) + ' dias): cadência de retomada' : '',
      Vars_JSON: JSON.stringify({
        produtos: produtosCarrinho_(cart),
        total: (cart.totalPrice && cart.totalPrice.formattedAmount) || formatarMoeda_(cart.totalPrice && cart.totalPrice.amount),
        link_checkout: cart.checkoutUrl || varsGlobais_().link_loja || '',
        data_carrinho: Utilities.formatDate(new Date(cart.createdDate), 'America/Sao_Paulo', 'dd/MM'),
        frase_b2b: corporativo
          ? 'Se a compra for pela empresa, emito orçamento formal e faturamento em CNPJ.'
          : '',
      }),
    });
    if (email) emCadencia[email.toLowerCase()] = true;
  });

  if (novos.length) {
    anexarLinhas_(ABA.CONTATOS, novos.map(o => CAB.CONTATOS.map(k => o[k] === undefined ? '' : o[k])));
    novos.forEach(o => log_(o.ID, 'INSCRITO', o.Cadencia + ' via Wix (' + o.Ref_Bling + ')'));
  }
  return novos.length;
}
