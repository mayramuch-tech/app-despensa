/**
 * Motor da cadência. Roda a cada hora (gatilho) ou pelo menu.
 * Ordem: Bling → inicializa novos → regras de parada (Agenda / e-mail) → executa passos vencidos.
 */

function processarCadencias() {
  const trava = LockService.getScriptLock();
  if (!trava.tryLock(5000)) { console.log('Execução anterior ainda em andamento.'); return; }
  try {
    _configCache = null;
    const agora = new Date();
    const resumo = { inscritos: 0, emails: 0, wpp: 0, tarefas: 0, encerrados: 0, erros: 0 };

    // 1) Bling
    if (cfgSim_('BLING_SINCRONIZAR') && props_().getProperty('BLING_REFRESH_TOKEN')) {
      try {
        const tabSync = lerTabela_(ABA.CONTATOS);
        resumo.inscritos = sincronizarBling_(tabSync);
        gravarTabela_(tabSync);
      } catch (e) {
        resumo.erros++;
        log_('', 'ERRO_BLING', e.message);
      }
    }

    const tab = lerTabela_(ABA.CONTATOS);
    const passos = lerPassos_();

    // 2) Inicializa linhas novas (leads manuais e inscrições do Bling)
    tab.linhas.forEach(l => {
      if (!l.Nome && !l.Email && !l.Telefone) return;
      if (l.Status) return;
      l.ID = l.ID || novoId_();
      l.Cadencia = l.Cadencia || cfg_().CADENCIA_LEAD_PADRAO;
      l.Origem = l.Origem || 'MANUAL';
      if (!passos[l.Cadencia]) { encerrar_(l, STATUS.ERRO, 'Cadência "' + l.Cadencia + '" não existe na aba Cadencias'); return; }
      l.Status = STATUS.ATIVO;
      l.Passo = 1;
      l.Inicio = agora;
      l.Proxima_Acao = somarDias_(agora, passos[l.Cadencia][0].Dia);
      l._sujo = true;
      log_(l.ID, 'INICIADO', l.Cadencia + ' | 1º passo em ' + l.Proxima_Acao);
    });

    // 3) Regras de parada
    const ativos = tab.linhas.filter(l => STATUS_FINAIS.indexOf(l.Status) < 0 && l.Status);
    if (ativos.length && cfgSim_('DETECTAR_REUNIAO')) {
      const convidados = convidadosDaAgenda_();
      ativos.forEach(l => {
        const r = l.Email && convidados[String(l.Email).toLowerCase()];
        if (r && r.inicio >= new Date(l.Inicio || 0).setHours(0, 0, 0, 0) - 86400000) {
          encerrar_(l, STATUS.REUNIAO, 'Evento na Agenda: "' + r.titulo + '" em ' + r.inicio);
          resumo.encerrados++;
        }
      });
    }
    if (cfgSim_('DETECTAR_RESPOSTA_EMAIL')) {
      ativos.filter(l => STATUS_FINAIS.indexOf(l.Status) < 0 && l.Email && Number(l.Passo) > 1).forEach(l => {
        if (respondeuEmail_(l.Email, new Date(l.Inicio))) {
          encerrar_(l, STATUS.RESPONDEU, 'Resposta recebida por e-mail');
          resumo.encerrados++;
        }
      });
    }

    // 4) Executa passos vencidos (só dentro do horário comercial)
    if (dentroDoHorario_(agora)) {
      const limite = Number(cfg_().LIMITE_EMAILS_DIA || 90);
      tab.linhas
        .filter(l => l.Status === STATUS.ATIVO && l.Proxima_Acao && new Date(l.Proxima_Acao) <= agora)
        .sort((a, b) => new Date(a.Proxima_Acao) - new Date(b.Proxima_Acao))
        .forEach(l => {
          const passo = (passos[l.Cadencia] || [])[Number(l.Passo) - 1];
          if (!passo) { encerrar_(l, STATUS.CONCLUIDO, 'Todos os passos executados'); return; }
          if (passo.Canal === CANAL.EMAIL && emailsEnviadosHoje_() >= limite) return; // fica para amanhã
          try {
            executarPasso_(l, passo, agora);
            if (passo.Canal === CANAL.EMAIL) resumo.emails++;
            else if (passo.Canal === CANAL.WHATSAPP) resumo.wpp++;
            else resumo.tarefas++;
          } catch (e) {
            resumo.erros++;
            l.Observacoes = 'Erro passo ' + l.Passo + ': ' + e.message;
            l._sujo = true;
            log_(l.ID, 'ERRO', e.message);
            // canal impossível para este contato (sem e-mail/telefone) → pula o passo
            if (/inválido|sem e-mail/i.test(e.message)) avancar_(l, passos, agora);
          }
        });
    }

    gravarTabela_(tab);
    log_('', 'CICLO', JSON.stringify(resumo));
    return resumo;
  } finally {
    trava.releaseLock();
  }
}

/** { CADENCIA: [passo1, passo2, ...] } ordenados pelo número do passo. */
function lerPassos_() {
  const mapa = {};
  lerTabela_(ABA.CADENCIAS).linhas.forEach(p => {
    if (!p.Cadencia || !p.Canal) return;
    p.Canal = String(p.Canal).trim().toUpperCase();
    (mapa[p.Cadencia] = mapa[p.Cadencia] || []).push(p);
  });
  Object.keys(mapa).forEach(k => mapa[k].sort((a, b) => Number(a.Passo) - Number(b.Passo)));
  return mapa;
}

function varsContato_(l) {
  let extra = {};
  try { extra = l.Vars_JSON ? JSON.parse(l.Vars_JSON) : {}; } catch (e) { /* JSON editado à mão */ }
  return Object.assign({}, varsGlobais_(), extra, {
    nome: l.Nome,
    primeiro_nome: primeiroNome_(l.Nome),
    empresa: l.Empresa || l.Nome,
    email: l.Email,
    telefone: l.Telefone,
  });
}

function executarPasso_(l, passo, agora) {
  const vars = varsContato_(l);
  const usaRastreio = /\{\{\s*rastreio\s*\}\}/i.test(passo.Assunto + passo.Mensagem);
  if (usaRastreio && vars.id_pedido_bling) vars.rastreio = rastreioPedido_(vars.id_pedido_bling) || 'em breve';
  const assunto = renderizar_(passo.Assunto, vars);
  const msg = renderizar_(passo.Mensagem, vars);
  const passos = lerPassosCache_();

  switch (passo.Canal) {
    case CANAL.EMAIL:
      if (!l.Email) throw new Error('Contato sem e-mail');
      enviarEmail_(l.Email, assunto, msg);
      contarEmailEnviado_();
      log_(l.ID, 'EMAIL', assunto);
      avancar_(l, passos, agora);
      break;

    case CANAL.WHATSAPP:
      enfileirarWhatsApp_(l, l.Passo, msg);
      l.Status = STATUS.AGUARDANDO_WPP; // avança quando o checkbox "Enviado" for marcado
      l.Ultima_Acao = agora;
      l._sujo = true;
      log_(l.ID, 'WPP_FILA', msg.substring(0, 120));
      break;

    case CANAL.LIGACAO:
    case CANAL.TAREFA: {
      const titulo = (passo.Canal === CANAL.LIGACAO ? '📞 ' : '✅ ') + (assunto || 'Contato: ' + l.Nome);
      criarTarefaAgenda_(l, titulo, msg, agora);
      log_(l.ID, 'AGENDA', titulo);
      avancar_(l, passos, agora);
      break;
    }

    default:
      throw new Error('Canal desconhecido: ' + passo.Canal);
  }
}

let _passosCache = null;
function lerPassosCache_() { return _passosCache || (_passosCache = lerPassos_()); }

/** Avança para o próximo passo. "Dia" é contado a partir do início; nunca antes de 4 h da ação atual. */
function avancar_(l, passos, agora) {
  const lista = passos[l.Cadencia] || [];
  const prox = Number(l.Passo) + 1;
  l.Ultima_Acao = agora;
  l._sujo = true;
  if (prox > lista.length) { encerrar_(l, STATUS.CONCLUIDO, 'Cadência finalizada'); return; }
  const previsto = somarDias_(new Date(l.Inicio || agora), lista[prox - 1].Dia);
  const minimo = new Date(agora.getTime() + 4 * 3600000);
  l.Passo = prox;
  l.Status = STATUS.ATIVO;
  l.Proxima_Acao = previsto > minimo ? previsto : minimo;
}

function encerrar_(l, status, motivo) {
  l.Status = status;
  l.Proxima_Acao = '';
  l.Observacoes = motivo;
  l._sujo = true;
  log_(l.ID, 'ENCERRADO_' + status, motivo);
}

/** Gatilho de edição: checkbox "Enviado" na Fila_WhatsApp → avança o contato. */
function aoEditar(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== ABA.FILA_WPP || e.range.getColumn() !== 9 || e.range.getRow() < 2) return;
  if (e.range.getValue() !== true) return;
  const contatoId = sh.getRange(e.range.getRow(), 2).getValue();
  const tab = lerTabela_(ABA.CONTATOS);
  const l = tab.linhas.find(x => x.ID === contatoId);
  if (!l || l.Status !== STATUS.AGUARDANDO_WPP) return;
  avancar_(l, lerPassos_(), new Date());
  gravarTabela_(tab);
  log_(l.ID, 'WPP_ENVIADO', 'Passo ' + (Number(l.Passo) - 1) + ' confirmado');
  sh.getRange(e.range.getRow(), 1, 1, 9).setBackground('#dcfce7');
}

/** Resumo diário (7h): WhatsApps pendentes, tarefas do dia e funil. */
function enviarResumoDiario() {
  _configCache = null;
  const para = String(cfg_().EMAIL_RESUMO_DIARIO || '').trim();
  if (!para) return;
  const contatos = lerTabela_(ABA.CONTATOS).linhas.filter(l => l.ID);
  const cont = {};
  contatos.forEach(l => { cont[l.Status] = (cont[l.Status] || 0) + 1; });
  const pendWpp = lerTabela_(ABA.FILA_WPP).linhas.filter(f => f.Enviado !== true && f.Contato_ID);

  const agenda = CalendarApp.getCalendarById(String(cfg_().CALENDARIO_ID || 'primary')) || CalendarApp.getDefaultCalendar();
  const ini = new Date(); ini.setHours(0, 0, 0, 0);
  const fim = new Date(ini.getTime() + 86400000);
  const eventos = agenda.getEvents(ini, fim);

  let html = '<h2>☀️ Cadência — resumo de ' + Utilities.formatDate(ini, 'America/Sao_Paulo', 'dd/MM/yyyy') + '</h2>';
  html += '<h3>Funil</h3><table border="1" cellpadding="4" style="border-collapse:collapse">' +
    Object.keys(cont).map(k => '<tr><td>' + k + '</td><td><b>' + cont[k] + '</b></td></tr>').join('') + '</table>';
  html += '<h3>💬 WhatsApp pendentes (' + pendWpp.length + ')</h3><ul>' +
    pendWpp.slice(0, 30).map(f => '<li>' + escaparHtml_(f.Nome) + ' — ' + f.Cadencia + ' passo ' + f.Passo +
      ' — <a href="' + linkWhatsApp_(f.Telefone, f.Mensagem) + '">abrir</a></li>').join('') + '</ul>';
  html += '<h3>📅 Agenda de hoje (' + eventos.length + ')</h3><ul>' +
    eventos.map(ev => '<li>' + Utilities.formatDate(ev.getStartTime(), 'America/Sao_Paulo', 'HH:mm') + ' — ' +
      escaparHtml_(ev.getTitle()) + '</li>').join('') + '</ul>';
  html += '<p><a href="' + SpreadsheetApp.getActive().getUrl() + '">Abrir painel</a></p>';

  GmailApp.sendEmail(para, 'Cadência | resumo do dia', 'Abra este e-mail em HTML.', { htmlBody: html, name: 'Cadência' });
}
