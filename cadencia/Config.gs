/**
 * CADÊNCIA INTEGRADA — Google Agenda × WhatsApp Business × Bling × E-mail UOL
 * Configuração, nomes de abas e estrutura da planilha-painel.
 */

const ABA = {
  CONFIG: 'Config',
  CADENCIAS: 'Cadencias',
  CONTATOS: 'Contatos',
  FILA_WPP: 'Fila_WhatsApp',
  LOG: 'Log',
  SITUACOES: 'Situacoes_Bling',
};

const CAB = {
  CONTATOS: ['ID', 'Cadencia', 'Status', 'Passo', 'Proxima_Acao', 'Nome', 'Empresa', 'Email', 'Telefone',
             'Origem', 'Ref_Bling', 'Situacao_Bling', 'Inicio', 'Ultima_Acao', 'Observacoes', 'Vars_JSON'],
  CADENCIAS: ['Cadencia', 'Passo', 'Dia', 'Canal', 'Assunto', 'Mensagem'],
  FILA_WPP: ['Data', 'Contato_ID', 'Nome', 'Telefone', 'Cadencia', 'Passo', 'Mensagem', 'Abrir_WhatsApp', 'Enviado'],
  LOG: ['DataHora', 'Contato_ID', 'Tipo', 'Detalhe'],
  SITUACOES: ['Modulo_ID', 'Modulo', 'Situacao_ID', 'Situacao'],
};

/** Status possíveis de um contato na cadência. */
const STATUS = {
  ATIVO: 'ATIVO',
  AGUARDANDO_WPP: 'AGUARDANDO_WPP',
  PAUSADO: 'PAUSADO',
  CONCLUIDO: 'CONCLUIDO',
  RESPONDEU: 'RESPONDEU',
  REUNIAO: 'REUNIAO_AGENDADA',
  GANHO: 'GANHO',
  PERDIDO: 'PERDIDO',
  CANCELADO: 'CANCELADO',
  ERRO: 'ERRO',
};

/** Status que encerram a cadência (o motor ignora esses contatos). */
const STATUS_FINAIS = [STATUS.PAUSADO, STATUS.CONCLUIDO, STATUS.RESPONDEU, STATUS.REUNIAO,
                       STATUS.GANHO, STATUS.PERDIDO, STATUS.CANCELADO, STATUS.ERRO];

const CANAL = { EMAIL: 'EMAIL', WHATSAPP: 'WHATSAPP', LIGACAO: 'LIGACAO', TAREFA: 'TAREFA' };

/** Parâmetros padrão. Chave | valor | descrição — gravados na aba Config na instalação. */
const CONFIG_PADRAO = [
  ['EMAIL_REMETENTE', 'contato@seudominio.com.br', 'Endereço UOL Host. Precisa estar cadastrado no Gmail em "Enviar e-mail como" (ver README).'],
  ['NOME_REMETENTE', 'Mayra Mucha | OxHydro', 'Nome exibido no remetente.'],
  ['ASSINATURA_HTML', '<p>--<br><b>Engª Mayra Mucha</b><br>Diretora Técnica (CTO) – OxHydro Saneamento</p>', 'Assinatura anexada a todo e-mail.'],
  ['EMAIL_RESUMO_DIARIO', '', 'E-mail que recebe o resumo das 7h (vazio = não envia).'],
  ['LIMITE_EMAILS_DIA', 90, 'Trava de segurança. Gmail gratuito: 100/dia; Workspace: 1.500/dia.'],
  ['HORA_INICIO', 8, 'Envios só a partir desta hora.'],
  ['HORA_FIM', 18, 'Envios só até esta hora.'],
  ['SO_DIAS_UTEIS', 'SIM', 'SIM = conta "Dia" da cadência em dias úteis e não envia no fim de semana.'],
  ['CALENDARIO_ID', 'primary', 'Agenda onde as tarefas/ligações são criadas e onde as reuniões são detectadas.'],
  ['HORA_TAREFAS', 9, 'Hora das tarefas/ligações criadas na Agenda.'],
  ['DURACAO_TAREFA_MIN', 15, 'Duração (min) do bloco de tarefa na Agenda.'],
  ['DETECTAR_REUNIAO', 'SIM', 'SIM = encerra a cadência se o contato for convidado de um evento na Agenda.'],
  ['DETECTAR_RESPOSTA_EMAIL', 'NAO', 'SIM = encerra se o contato responder (requer caixa UOL importada no Gmail via POP).'],
  ['BLING_SINCRONIZAR', 'SIM', 'SIM = importa pedidos e propostas do Bling a cada execução.'],
  ['BLING_JANELA_DIAS', 15, 'Quantos dias para trás o Bling é varrido em cada sincronização.'],
  ['BLING_PEDIDO_SITUACOES_ENTRADA', '6,9,15', 'IDs de situação do pedido que inscrevem na cadência (padrão Bling: 6 Em aberto, 9 Atendido, 15 Em andamento).'],
  ['BLING_PEDIDO_SITUACOES_PARADA', '12', 'IDs que encerram a cadência (padrão: 12 Cancelado).'],
  ['BLING_PROPOSTA_SITUACOES_GANHO', 'Aprovado,Concluído', 'Situações de proposta que encerram como GANHO.'],
  ['BLING_PROPOSTA_SITUACOES_PERDA', 'Não aprovado', 'Situações de proposta que encerram como PERDIDO.'],
  ['BLING_PROPOSTA_SITUACOES_IGNORAR', 'Rascunho', 'Situações de proposta que não entram na cadência.'],
  ['CADENCIA_PEDIDO', 'POS_VENDA', 'Cadência aplicada a pedidos do Bling.'],
  ['CADENCIA_PROPOSTA', 'PROPOSTA', 'Cadência aplicada a propostas comerciais do Bling.'],
  ['CADENCIA_LEAD_PADRAO', 'PROSPECCAO_B2B', 'Cadência para linhas novas em Contatos sem cadência preenchida.'],
  ['VAR_LINK_AVALIACAO', 'https://seusite.com.br/avaliar', 'Variável {{link_avaliacao}} nos modelos.'],
  ['VAR_LINK_LOJA', 'https://seusite.com.br', 'Variável {{link_loja}} nos modelos.'],
  ['VAR_LINK_AGENDA', 'https://calendar.app.google/seu-link', 'Variável {{link_agenda}} (página de agendamento do Google Agenda).'],
  ['VAR_ASSINATURA_WPP', 'Mayra | OxHydro', 'Variável {{assinatura_wpp}}.'],
];

let _configCache = null;

/** Lê a aba Config como objeto { CHAVE: valor }. */
function cfg_() {
  if (_configCache) return _configCache;
  const sh = aba_(ABA.CONFIG);
  const valores = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const c = {};
  CONFIG_PADRAO.forEach(([k, v]) => { c[k] = v; });
  valores.forEach(([k, v]) => { if (k) c[String(k).trim()] = v; });
  _configCache = c;
  return c;
}

function cfgSim_(chave) {
  return String(cfg_()[chave]).trim().toUpperCase() === 'SIM';
}

function cfgLista_(chave) {
  return String(cfg_()[chave] || '').split(',').map(s => normaliza_(s)).filter(String);
}

/** Variáveis globais de modelo: toda chave VAR_XYZ vira {{xyz}}. */
function varsGlobais_() {
  const c = cfg_(), v = {};
  Object.keys(c).forEach(k => { if (k.indexOf('VAR_') === 0) v[k.substring(4).toLowerCase()] = c[k]; });
  return v;
}

/** Cria/garante todas as abas, cabeçalhos e dados iniciais. Idempotente. */
function instalarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  const config = garanteAba_(ABA.CONFIG, ['Chave', 'Valor', 'Descricao']);
  const existentes = config.getLastRow() > 1
    ? config.getRange(2, 1, config.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const novas = CONFIG_PADRAO.filter(r => existentes.indexOf(r[0]) < 0);
  if (novas.length) config.getRange(config.getLastRow() + 1, 1, novas.length, 3).setValues(novas);
  config.setColumnWidth(1, 260).setColumnWidth(2, 320).setColumnWidth(3, 520);

  const cad = garanteAba_(ABA.CADENCIAS, CAB.CADENCIAS);
  if (cad.getLastRow() < 2) {
    const modelos = cadenciasModelo_();
    cad.getRange(2, 1, modelos.length, CAB.CADENCIAS.length).setValues(modelos);
  }
  cad.setColumnWidth(5, 280).setColumnWidth(6, 520);
  cad.getRange('F:F').setWrap(true);
  cad.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CANAL), true).build());

  const ct = garanteAba_(ABA.CONTATOS, CAB.CONTATOS);
  ct.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(STATUS).map(k => STATUS[k]), true).setAllowInvalid(true).build());
  ct.getRange('E2:E').setNumberFormat('dd/mm/yyyy hh:mm');
  ct.getRange('M2:N').setNumberFormat('dd/mm/yyyy hh:mm');
  ct.getRange('I2:I').setNumberFormat('@');

  const fila = garanteAba_(ABA.FILA_WPP, CAB.FILA_WPP);
  fila.setColumnWidth(7, 420);
  fila.getRange('A2:A').setNumberFormat('dd/mm/yyyy hh:mm');

  garanteAba_(ABA.LOG, CAB.LOG).getRange('A2:A').setNumberFormat('dd/mm/yyyy hh:mm:ss');
  garanteAba_(ABA.SITUACOES, CAB.SITUACOES);

  _configCache = null;
  SpreadsheetApp.getActive().toast('Planilha pronta. Próximo passo: Cadência ▸ Configurar Bling.', 'Cadência', 8);
}

function garanteAba_(nome, cabecalho) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Cadências-modelo. Edite à vontade na aba Cadencias; "Dia" é contado a partir da inscrição. */
function cadenciasModelo_() {
  return [
    // ── PÓS-VENDA E-COMMERCE (pedido Bling) ──────────────────────────────
    ['POS_VENDA', 1, 0, 'EMAIL', 'Recebemos seu pedido nº {{numero_pedido}} ✅',
      'Olá, {{primeiro_nome}}!\n\nObrigada pela compra. Seu pedido nº {{numero_pedido}} (R$ {{total}}) foi registrado e já está em preparação.\n\nAssim que for despachado, você recebe o código de rastreio.\n\nQualquer dúvida, é só responder este e-mail.'],
    ['POS_VENDA', 2, 1, 'WHATSAPP', '',
      'Oi, {{primeiro_nome}}! Aqui é da OxHydro 😊 Recebemos seu pedido nº {{numero_pedido}}. Se precisar de orientação de uso ou diluição dos produtos Oxy, me chama por aqui. {{assinatura_wpp}}'],
    ['POS_VENDA', 3, 7, 'EMAIL', 'Como aproveitar ao máximo seu produto Oxy',
      'Olá, {{primeiro_nome}}!\n\nSeu pedido já deve ter chegado. Separamos as orientações principais:\n\n• Siga sempre a diluição indicada no rótulo.\n• Armazene em local fresco, ao abrigo da luz.\n• Mantenha a embalagem bem fechada após o uso.\n\nSe tiver qualquer dúvida técnica, responda este e-mail que a nossa equipe técnica orienta você.'],
    ['POS_VENDA', 4, 15, 'WHATSAPP', '',
      'Oi, {{primeiro_nome}}! Tudo certo com seu pedido? Sua opinião ajuda muito outros clientes: {{link_avaliacao}} — leva menos de 1 minuto. Obrigada! {{assinatura_wpp}}'],
    ['POS_VENDA', 5, 45, 'EMAIL', '{{primeiro_nome}}, hora de repor?',
      'Olá, {{primeiro_nome}}!\n\nPelo consumo médio, seu estoque pode estar no fim. Para não ficar sem, você pode repor direto na loja: {{link_loja}}\n\nConte com a gente!'],

    // ── PROSPECÇÃO B2B (leads manuais) ───────────────────────────────────
    ['PROSPECCAO_B2B', 1, 0, 'EMAIL', '{{empresa}} | tratamento e desinfecção de água com dióxido de cloro',
      'Olá, {{primeiro_nome}}, tudo bem?\n\nSou a Mayra, Diretora Técnica da OxHydro Saneamento. Atuamos com projeto e operação de sistemas de tratamento de água e efluentes e com desinfecção à base de dióxido de cloro estabilizado (linha Oxy).\n\nGostaria de entender o cenário da {{empresa}} e avaliar se faz sentido uma conversa técnica de 20 minutos.\n\nSe preferir, escolha um horário aqui: {{link_agenda}}'],
    ['PROSPECCAO_B2B', 2, 2, 'WHATSAPP', '',
      'Olá, {{primeiro_nome}}! Aqui é a Mayra, da OxHydro. Enviei um e-mail sobre tratamento/desinfecção de água para a {{empresa}}. Faz sentido conversarmos 20 min esta semana? {{link_agenda}}'],
    ['PROSPECCAO_B2B', 3, 4, 'LIGACAO', 'Ligar para {{nome}} ({{empresa}})',
      'Roteiro: 1) apresentar OxHydro; 2) entender sistema atual (fonte, vazão, desinfecção, problemas); 3) propor visita/diagnóstico técnico. Tel: {{telefone}}'],
    ['PROSPECCAO_B2B', 4, 7, 'EMAIL', 'Como reduzimos risco e custo operacional em sistemas como o da {{empresa}}',
      'Olá, {{primeiro_nome}}!\n\nRetomando meu contato: em projetos semelhantes, o diagnóstico técnico inicial costuma apontar ganhos rápidos em segurança sanitária e em custo de operação.\n\nPosso apresentar um caso aplicado em 20 minutos? {{link_agenda}}'],
    ['PROSPECCAO_B2B', 5, 11, 'WHATSAPP', '',
      '{{primeiro_nome}}, sigo à disposição para um diagnóstico técnico sem compromisso na {{empresa}}. Qual o melhor dia para você? {{assinatura_wpp}}'],
    ['PROSPECCAO_B2B', 6, 14, 'EMAIL', 'Encerro por aqui, {{primeiro_nome}}?',
      'Olá, {{primeiro_nome}}!\n\nComo não tivemos retorno, vou pausar meus contatos para não ser inconveniente. Se o tema voltar à pauta da {{empresa}}, é só responder este e-mail.\n\nUm abraço!'],

    // ── PROPOSTAS COMERCIAIS (Bling) ─────────────────────────────────────
    ['PROPOSTA', 1, 1, 'EMAIL', 'Proposta nº {{numero_proposta}} | {{empresa}}',
      'Olá, {{primeiro_nome}}!\n\nConfirmo o envio da proposta nº {{numero_proposta}}, no valor de R$ {{total}}. Fico à disposição para esclarecer qualquer ponto técnico ou comercial.'],
    ['PROPOSTA', 2, 3, 'WHATSAPP', '',
      'Oi, {{primeiro_nome}}! Conseguiu avaliar a proposta nº {{numero_proposta}}? Posso tirar dúvidas por aqui ou numa chamada rápida. {{assinatura_wpp}}'],
    ['PROPOSTA', 3, 5, 'LIGACAO', 'Follow-up proposta {{numero_proposta}} – {{nome}}',
      'Objetivo: entender objeções (escopo, prazo, preço) e próximo passo de decisão. Tel: {{telefone}}'],
    ['PROPOSTA', 4, 8, 'EMAIL', 'Ajustes na proposta nº {{numero_proposta}}?',
      'Olá, {{primeiro_nome}}!\n\nSe houver necessidade de ajustar escopo, prazo ou condições de pagamento da proposta nº {{numero_proposta}}, me diga que reviso.'],
    ['PROPOSTA', 5, 15, 'EMAIL', 'Validade da proposta nº {{numero_proposta}}',
      'Olá, {{primeiro_nome}}!\n\nLembrando que a proposta nº {{numero_proposta}} está chegando ao fim da validade. Quer que eu reserve as condições atuais?'],
  ];
}
