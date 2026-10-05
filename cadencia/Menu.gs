/** Menu da planilha e instalação de gatilhos. */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Cadência')
    .addItem('▶️ Processar agora', 'menuProcessar')
    .addSeparator()
    .addItem('⏸️ Pausar contato(s) selecionado(s)', 'pausarSelecionados')
    .addItem('⏯️ Retomar contato(s) selecionado(s)', 'retomarSelecionados')
    .addItem('🏆 Marcar selecionado(s) como GANHO', 'ganharSelecionados')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔌 Bling')
      .addItem('Configurar credenciais', 'configurarBling')
      .addItem('Autorizar / reconectar', 'autorizarBling')
      .addItem('Testar conexão', 'testarBling')
      .addItem('Listar situações (IDs)', 'listarSituacoesBling'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠️ Instalação')
      .addItem('1. Instalar planilha', 'instalarPlanilha')
      .addItem('2. Testar e-mail UOL', 'testarEmail')
      .addItem('3. Ativar automação (gatilhos)', 'instalarGatilhos')
      .addItem('Desativar automação', 'removerGatilhos'))
    .addToUi();
}

function menuProcessar() {
  const r = processarCadencias();
  if (!r) return;
  const fora = dentroDoHorario_(new Date()) ? '' : '\n\n⚠️ Fora do horário de envio: só sincronizou e verificou paradas.';
  SpreadsheetApp.getUi().alert('Ciclo concluído:\n\n' +
    '• Novos do Bling: ' + r.inscritos + '\n• E-mails: ' + r.emails + '\n• WhatsApp na fila: ' + r.wpp +
    '\n• Tarefas na Agenda: ' + r.tarefas + '\n• Encerrados: ' + r.encerrados + '\n• Erros: ' + r.erros + fora);
}

function instalarGatilhos() {
  removerGatilhos(true);
  ScriptApp.newTrigger('processarCadencias').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('enviarResumoDiario').timeBased().atHour(7).everyDays(1).inTimezone('America/Sao_Paulo').create();
  ScriptApp.newTrigger('aoEditar').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getUi().alert('✅ Automação ativa:\n• Motor a cada 1 h\n• Resumo diário às 7h\n• Checkbox da Fila_WhatsApp avança o contato');
}

function removerGatilhos(silencioso) {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  if (silencioso !== true) SpreadsheetApp.getUi().alert('Automação desativada.');
}

function testarEmail() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Testar e-mail', 'Enviar teste para qual endereço?', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    enviarEmail_(r.getResponseText().trim(), 'Teste da cadência integrada',
      'Se você recebeu este e-mail com o remetente ' + cfg_().EMAIL_REMETENTE + ', o envio pelo UOL está funcionando. ✅');
    ui.alert('✅ Enviado. Confira a caixa de entrada (e o remetente).');
  } catch (e) {
    ui.alert('❌ ' + e.message);
  }
}

/** Altera o status das linhas selecionadas na aba Contatos. */
function alterarSelecionados_(fn) {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== ABA.CONTATOS) { SpreadsheetApp.getUi().alert('Selecione linhas na aba Contatos.'); return; }
  const tab = lerTabela_(ABA.CONTATOS);
  const sel = sh.getActiveRangeList().getRanges();
  let n = 0;
  sel.forEach(r => {
    for (let i = r.getRow(); i < r.getRow() + r.getNumRows(); i++) {
      const l = tab.linhas.find(x => x._linha === i);
      if (l && l.ID) { fn(l); n++; }
    }
  });
  gravarTabela_(tab);
  SpreadsheetApp.getActive().toast(n + ' contato(s) atualizado(s).', 'Cadência', 4);
}

function pausarSelecionados() {
  alterarSelecionados_(l => encerrar_(l, STATUS.PAUSADO, 'Pausado manualmente'));
}

function ganharSelecionados() {
  alterarSelecionados_(l => encerrar_(l, STATUS.GANHO, 'Marcado como ganho manualmente'));
}

function retomarSelecionados() {
  alterarSelecionados_(l => {
    if (l.Status !== STATUS.PAUSADO && l.Status !== STATUS.ERRO) return;
    l._sujo = true;
    if (!l.Passo) { l.Status = ''; return; } // erro na inicialização → reinicia no próximo ciclo
    l.Status = STATUS.ATIVO;
    l.Proxima_Acao = new Date();
    l.Observacoes = 'Retomado manualmente';
    l._sujo = true;
    log_(l.ID, 'RETOMADO', 'Passo ' + l.Passo);
  });
}
