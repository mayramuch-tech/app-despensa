/**
 * Canais de execução de cada passo:
 *  • EMAIL    → envio pelo Gmail usando o endereço UOL Host como remetente ("Enviar e-mail como").
 *  • WHATSAPP → fila semiautomática com link wa.me (1 clique abre a conversa com o texto pronto).
 *  • LIGACAO / TAREFA → evento no Google Agenda.
 */

/** Envia e-mail com remetente UOL. Lança erro se o alias não estiver configurado no Gmail. */
function enviarEmail_(para, assunto, corpoTexto) {
  const c = cfg_();
  const remetente = String(c.EMAIL_REMETENTE).trim();
  const aliases = GmailApp.getAliases().map(a => a.toLowerCase());
  const minhaConta = Session.getActiveUser().getEmail().toLowerCase();
  if (remetente.toLowerCase() !== minhaConta && aliases.indexOf(remetente.toLowerCase()) < 0) {
    throw new Error('O endereço ' + remetente + ' não está em "Enviar e-mail como" no Gmail. Veja o README, passo 2.');
  }
  const html = textoParaHtml_(corpoTexto) + String(c.ASSINATURA_HTML || '');
  const textoAssinatura = String(c.ASSINATURA_HTML || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
  GmailApp.sendEmail(para, assunto, corpoTexto + '\n\n' + textoAssinatura, {
    from: remetente,
    name: String(c.NOME_REMETENTE || ''),
    replyTo: remetente,
    htmlBody: html,
  });
}

/** Quantos e-mails a cadência já enviou hoje (controle de cota). */
function emailsEnviadosHoje_() {
  const p = PropertiesService.getScriptProperties();
  const hoje = dataISO_(new Date());
  return p.getProperty('EMAILS_DIA') === hoje ? Number(p.getProperty('EMAILS_QTD') || 0) : 0;
}

function contarEmailEnviado_() {
  const p = PropertiesService.getScriptProperties();
  p.setProperties({ EMAILS_DIA: dataISO_(new Date()), EMAILS_QTD: String(emailsEnviadosHoje_() + 1) });
}

/** Coloca a mensagem na Fila_WhatsApp com link pronto e checkbox "Enviado". */
function enfileirarWhatsApp_(contato, passo, mensagem) {
  const link = linkWhatsApp_(contato.Telefone, mensagem);
  if (!link) throw new Error('Telefone inválido para WhatsApp: "' + contato.Telefone + '"');
  const sh = aba_(ABA.FILA_WPP);
  const linha = sh.getLastRow() + 1;
  sh.getRange(linha, 1, 1, 9).setValues([[
    new Date(), contato.ID, contato.Nome, telefoneWpp_(contato.Telefone), contato.Cadencia, passo,
    mensagem, '=HYPERLINK("' + link.replace(/"/g, '""') + '";"💬 Abrir")', false,
  ]]);
  sh.getRange(linha, 9).insertCheckboxes();
  sh.getRange(linha, 7).setWrap(true);
}

/** Cria a tarefa/ligação na Agenda no dia do passo. */
function criarTarefaAgenda_(contato, titulo, descricao, quando) {
  const c = cfg_();
  const agenda = CalendarApp.getCalendarById(String(c.CALENDARIO_ID || 'primary')) || CalendarApp.getDefaultCalendar();
  const inicio = new Date(quando.getTime());
  inicio.setHours(Number(c.HORA_TAREFAS || 9), 0, 0, 0);
  const agora = new Date();
  if (inicio < agora) { // já passou do horário de hoje → próxima hora cheia
    inicio.setTime(agora.getTime());
    inicio.setHours(agora.getHours() + 1, 0, 0, 0);
  }
  const fim = new Date(inicio.getTime() + Number(c.DURACAO_TAREFA_MIN || 15) * 60000);
  const ev = agenda.createEvent(titulo, inicio, fim, {
    description: descricao + '\n\n[cadencia:' + contato.ID + ']' +
      (contato.Email ? '\nE-mail: ' + contato.Email : '') +
      (contato.Telefone ? '\nWhatsApp: ' + (linkWhatsApp_(contato.Telefone, '') || contato.Telefone) : ''),
  });
  ev.addPopupReminder(10);
  return ev;
}

/**
 * E-mails de convidados de eventos na Agenda (últimos 30 dias até +60 dias).
 * Eventos criados pela própria cadência (tarefas) são ignorados.
 */
function convidadosDaAgenda_() {
  const agenda = CalendarApp.getCalendarById(String(cfg_().CALENDARIO_ID || 'primary')) || CalendarApp.getDefaultCalendar();
  const ini = new Date(); ini.setDate(ini.getDate() - 30);
  const fim = new Date(); fim.setDate(fim.getDate() + 60);
  const mapa = {};
  agenda.getEvents(ini, fim).forEach(ev => {
    if (String(ev.getDescription() || '').indexOf('[cadencia:') >= 0) return;
    ev.getGuestList(true).forEach(g => {
      const e = g.getEmail().toLowerCase();
      if (!mapa[e] || ev.getStartTime() < mapa[e].inicio) mapa[e] = { inicio: ev.getStartTime(), titulo: ev.getTitle() };
    });
  });
  return mapa;
}

/** O contato respondeu por e-mail depois da inscrição? (exige caixa UOL importada no Gmail) */
function respondeuEmail_(email, desde) {
  if (!email) return false;
  const q = 'from:' + email + ' after:' + Utilities.formatDate(desde, 'America/Sao_Paulo', 'yyyy/MM/dd');
  return GmailApp.search(q, 0, 1).length > 0;
}
