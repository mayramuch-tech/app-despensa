/**
 * Gera cadencia/Manual_Cadencia_Integrada.docx.
 * Uso (de dentro de cadencia/manual):  node gerar_manual.js
 * Parâmetros e cadências são lidos direto de ../Config.gs, para o manual nunca divergir do código.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const D = require('docx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, Footer, Header, PageNumber, TableOfContents, PageBreak } = D;

const RAIZ = path.join(__dirname, '..');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(RAIZ, 'Config.gs'), 'utf8') +
  '\nthis.CP = CONFIG_PADRAO; this.CM = cadenciasModelo_();', ctx);
const CONFIG = ctx.CP, MODELOS = ctx.CM;

const AZUL = '0F172A', LARANJA = 'EA580C', CINZA = 'F1F5F9', BORDA = 'CBD5E1';
const LARG = 9638; // A4 com margens de 2 cm

const t = (text, o = {}) => new TextRun({ text, ...o });
const p = (conteudo, o = {}) => new Paragraph({
  spacing: { after: 120, line: 276 }, ...o,
  children: (Array.isArray(conteudo) ? conteudo : [conteudo]).map(c => typeof c === 'string' ? t(c) : c),
});
const h1 = s => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(s)] });
const h2 = s => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(s)] });
const num = (conteudo, ref = 'num') => p(conteudo, { numbering: { reference: ref, level: 0 }, spacing: { after: 60 } });
const b = s => t(s, { bold: true });
const code = s => t(s, { font: 'Consolas', size: 19, color: '9A3412' });
const esp = () => p('', { spacing: { after: 60 } });

function nota(titulo, texto, cor = 'FFF7ED', borda = LARANJA) {
  return new Paragraph({
    spacing: { before: 120, after: 160 },
    shading: { type: ShadingType.CLEAR, fill: cor, color: 'auto' },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: borda, space: 6 } },
    indent: { left: 120, right: 120 },
    children: [b(titulo + ' '), ...(Array.isArray(texto) ? texto : [t(texto)])],
  });
}

function tabela(cab, linhas, pesos) {
  const soma = pesos.reduce((a, x) => a + x, 0);
  const larg = pesos.map(x => Math.floor(LARG * x / soma));
  larg[larg.length - 1] += LARG - larg.reduce((a, x) => a + x, 0);
  const borda = { style: BorderStyle.SINGLE, size: 4, color: BORDA };
  const bordas = { top: borda, bottom: borda, left: borda, right: borda };
  const cel = (conteudo, i, cabecalho, zebra) => new TableCell({
    width: { size: larg[i], type: WidthType.DXA },
    borders: bordas,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: cabecalho ? { type: ShadingType.CLEAR, fill: AZUL, color: 'auto' }
      : zebra ? { type: ShadingType.CLEAR, fill: CINZA, color: 'auto' } : undefined,
    children: String(conteudo).split('\n').map(l => new Paragraph({
      spacing: { after: 40 },
      children: [t(l, { size: 18, bold: cabecalho, color: cabecalho ? 'FFFFFF' : undefined })],
    })),
  });
  return new Table({
    width: { size: LARG, type: WidthType.DXA },
    columnWidths: larg,
    rows: [
      new TableRow({ tableHeader: true, children: cab.map((c, i) => cel(c, i, true)) }),
      ...linhas.map((l, r) => new TableRow({ cantSplit: true, children: l.map((c, i) => cel(c, i, false, r % 2 === 1)) })),
    ],
  });
}

const corpo = [];

// ── Capa e sumário ──
corpo.push(
  p('', { spacing: { before: 2400 } }),
  p(t('MANUAL DE INSTALAÇÃO E OPERAÇÃO', { size: 22, bold: true, color: LARANJA })),
  p(t('Cadência Integrada', { size: 60, bold: true, color: AZUL }), { spacing: { after: 60 } }),
  p(t('Google Agenda × WhatsApp Business × Bling × Wix × E-mail UOL', { size: 28, color: '475569' }), {
    spacing: { after: 600 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: LARANJA, space: 8 } },
  }),
  p([b('Plataforma: '), t('Google Apps Script + Google Sheets (painel)')]),
  p([b('Fluxos: '), t('Pós-venda · Carrinho abandonado · Prospecção B2B · Propostas comerciais')]),
  p([b('Responsável técnica: '), t('Engª Mayra Mucha – Diretora Técnica (CTO), OxHydro Saneamento')]),
  p([b('Versão: '), t('1.1 – novembro/2026 (inclui carrinhos abandonados do Wix)')]),
  p([b('Código-fonte: '), t('repositório app-despensa, pasta cadencia/')]),
  new Paragraph({ children: [new PageBreak()] }),
  h1('Sumário'),
  new TableOfContents('Sumário', { hyperlink: true, headingStyleRange: '1-2' }),
  p(t('Se o sumário aparecer vazio, clique com o botão direito sobre ele e escolha "Atualizar campo".', { italics: true, size: 18, color: '64748B' })),
  new Paragraph({ children: [new PageBreak()] }),
);

// ── 1. Visão geral ──
corpo.push(
  h1('1. Visão geral'),
  p('A Cadência Integrada automatiza o relacionamento com clientes e leads, reunindo cinco ferramentas num só fluxo. Uma planilha Google funciona como painel de controle. O motor, escrito em Google Apps Script, roda a cada hora: busca dados novos no Bling e no Wix, verifica a Agenda, envia os e-mails e prepara as mensagens de WhatsApp.'),
  h2('1.1 Fluxos atendidos'),
  tabela(['Cadência', 'Origem', 'Entra quando', 'Encerra quando'], [
    ['POS_VENDA', 'Bling – pedidos de venda', 'Pedido nas situações de entrada (padrão 6, 9 ou 15)', 'Pedido cancelado (12) ou fim dos passos'],
    ['CARRINHO_ABANDONADO', 'Wix – carrinhos abandonados', 'Carrinho aberto há mais de 60 min e menos de 7 dias', 'Carrinho recuperado (GANHO) ou fim dos passos'],
    ['CARRINHO_RETOMADA', 'Wix – carrinhos antigos', 'Carrinho aberto há mais de 7 dias (até 60 dias)', 'Carrinho recuperado (GANHO) ou fim dos passos'],
    ['PROSPECCAO_B2B', 'Linha digitada na aba Contatos', 'Linha nova com Status vazio', 'Reunião na Agenda, resposta por e-mail, pausa ou fim dos passos'],
    ['PROPOSTA', 'Bling – propostas comerciais', 'Proposta nova (exceto Rascunho)', 'Aprovada/Concluída = GANHO; Não aprovada = PERDIDO'],
  ], [2.4, 2.4, 3, 3.2]),
  esp(),
  h2('1.2 Papel de cada ferramenta'),
  tabela(['Ferramenta', 'O que o programa faz', 'Nível de automação'], [
    ['Bling', 'Importa pedidos e propostas, lê os dados do contato e acompanha a mudança de situação', 'Automático'],
    ['Wix', 'Importa carrinhos abandonados (produtos, valor, link do checkout) e detecta quando o cliente conclui a compra', 'Automático'],
    ['E-mail UOL', 'Envia pelo Gmail com o endereço UOL Host como remetente', 'Automático'],
    ['WhatsApp Business', 'Monta a mensagem numa fila com link wa.me. A pessoa clica, revisa, envia e marca "Enviado"', 'Semiautomático'],
    ['Google Agenda', 'Cria as tarefas de ligação com roteiro e encerra a cadência quando há reunião com o contato', 'Automático'],
  ], [2, 5.5, 2]),
  nota('Por que o WhatsApp é semiautomático?', 'Esse modelo não tem custo nem exige aprovação da Meta, e não corre risco de bloqueio do número. Se o volume crescer, a fila pode ser trocada pela API oficial (WhatsApp Cloud API) sem mudar o restante do sistema.'),
  h2('1.3 Ciclo do motor (a cada 1 hora)'),
  num([b('Sincroniza o Bling: '), t('inscreve pedidos e propostas novos e atualiza a situação dos que já estão na cadência.')], 'ciclo'),
  num([b('Sincroniza o Wix: '), t('inscreve carrinhos abandonados e encerra os que foram recuperados.')], 'ciclo'),
  num([b('Inicializa contatos novos: '), t('linhas sem Status recebem ID, passo 1 e data da próxima ação.')], 'ciclo'),
  num([b('Aplica as regras de parada: '), t('reunião na Agenda, resposta por e-mail (opcional) e situação no Bling/Wix.')], 'ciclo'),
  num([b('Executa os passos vencidos: '), t('só dentro do horário comercial e, por padrão, só em dias úteis.')], 'ciclo'),
  num([b('Registra tudo na aba Log.')], 'ciclo'),
  p('Às 7h chega um resumo por e-mail com o funil, os WhatsApps pendentes e a agenda do dia.'),
);

// ── 2. Pré-requisitos ──
corpo.push(
  h1('2. Pré-requisitos'),
  tabela(['Item', 'Detalhe'], [
    ['Conta Google', 'A mesma conta Gmail/Workspace que vai enviar os e-mails e onde está a Agenda'],
    ['Caixa UOL Host', 'Endereço (ex.: contato@oxhydro.com.br) e senha da caixa'],
    ['Bling', 'Usuário com acesso à Central de Extensões para criar um aplicativo próprio'],
    ['Wix', 'Acesso de proprietário à conta Wix para gerar uma API key'],
    ['WhatsApp Business', 'Instalado no celular e conectado ao WhatsApp Web/Desktop no computador de operação'],
    ['Arquivos', 'Pasta cadencia/: appsscript.json, Config.gs, Util.gs, Canais.gs, Bling.gs, Wix.gs, Motor.gs, Menu.gs'],
  ], [2.2, 7.8]),
);

// ── 3. Instalação ──
corpo.push(
  h1('3. Instalação'),
  p('Tempo estimado: 40 minutos. Siga a ordem; cada etapa depende da anterior.'),
  h2('3.1 Planilha e script'),
  num('Crie uma planilha nova no Google Sheets com o nome "Cadência OxHydro".', 'i1'),
  num([t('Abra '), b('Extensões ▸ Apps Script'), t('.')], 'i1'),
  num('Crie os sete arquivos .gs com os mesmos nomes dos arquivos da pasta cadencia/ e cole o conteúdo de cada um.', 'i1'),
  num([t('Em '), b('Configurações do projeto'), t(', marque "Mostrar arquivo de manifesto appsscript.json" e substitua o conteúdo pelo do arquivo da pasta.')], 'i1'),
  num([t('Salve e recarregue a planilha. O menu '), b('⚙️ Cadência'), t(' aparece na barra.')], 'i1'),
  num([t('Rode '), b('⚙️ Cadência ▸ Instalação ▸ 1. Instalar planilha'), t(' e autorize as permissões solicitadas.')], 'i1'),
  nota('Já usa a versão 1.0?', 'Substitua os arquivos, crie o Wix.gs e rode "1. Instalar planilha" de novo. Ele só acrescenta as configurações e cadências novas; os textos que você já editou são preservados.', 'EFF6FF', '2563EB'),

  h2('3.2 E-mail UOL Host no Gmail'),
  p('O Apps Script não abre conexão SMTP direta. Por isso o envio sai pelo Gmail, com o endereço UOL como remetente.'),
  num([t('No Gmail, abra '), b('Configurações ▸ Ver todas ▸ Contas e importação ▸ Enviar e-mail como ▸ Adicionar outro endereço'), t('.')], 'i2'),
  num([t('Informe o nome e o e-mail UOL e '), b('desmarque "Tratar como alias"'), t('.')], 'i2'),
  num('Preencha os dados de SMTP da tabela abaixo.', 'i2'),
  num('Confirme pelo código que chega na caixa UOL.', 'i2'),
  tabela(['Campo', 'Valor'], [
    ['Servidor SMTP', 'smtps.uhserver.com'],
    ['Porta / segurança', '465 com SSL (alternativa: smtp.uhserver.com, 587, TLS)'],
    ['Usuário', 'E-mail completo'],
    ['Senha', 'Senha da caixa UOL'],
  ], [3, 7]),
  esp(),
  p([b('Opcional, para detectar respostas: '), t('em "Verificar o e-mail de outras contas", adicione POP3 pop.uhserver.com, porta 995, SSL. Depois mude DETECTAR_RESPOSTA_EMAIL para SIM na aba Config.')]),

  h2('3.3 Publicação como App da Web'),
  p('Essa publicação gera o endereço de retorno que o Bling usa para autorizar o acesso.'),
  num([t('No Apps Script, clique em '), b('Implantar ▸ Nova implantação ▸ Tipo: App da Web'), t('.')], 'i3'),
  num([t('Escolha '), b('Executar como: Eu'), t(' e '), b('Quem pode acessar: Somente eu'), t('.')], 'i3'),
  num([t('Copie a URL que termina em '), code('/exec'), t('.')], 'i3'),

  h2('3.4 Aplicativo no Bling'),
  num([t('No Bling, abra '), b('Central de Extensões ▸ Área do Integrador ▸ Criar aplicativo'), t(' (tipo: aplicativo próprio).')], 'i4'),
  num([t('Em '), b('Link de redirecionamento'), t(', cole a URL /exec.')], 'i4'),
  num('Marque os escopos de leitura: Pedidos de Venda, Propostas Comerciais, Contatos, Situações e Dados básicos da empresa.', 'i4'),
  num('Copie o Client ID e o Client Secret.', 'i4'),
  num([t('Na planilha, rode '), b('⚙️ Cadência ▸ Bling ▸ Configurar credenciais'), t(', cole os dois códigos e clique em "Autorizar no Bling".')], 'i4'),
  num([t('Rode '), b('Bling ▸ Testar conexão'), t('. O nome da empresa deve aparecer.')], 'i4'),
  num([t('Rode '), b('Bling ▸ Listar situações (IDs)'), t(', confira os códigos na aba Situacoes_Bling e ajuste as chaves BLING_* na aba Config.')], 'i4'),
  nota('Segurança:', 'Client ID, Client Secret, API key do Wix e tokens ficam guardados nas propriedades internas do script, nunca na planilha. O token do Bling é renovado sozinho. Se o Bling ficar mais de 30 dias sem sincronizar, rode "Autorizar / reconectar".'),

  h2('3.5 Loja Wix (carrinhos abandonados)'),
  num([t('Em '), b('manage.wix.com/account/api-keys'), t(', crie uma API key com permissão de leitura de pedidos do Wix Stores/eCommerce (Read Orders).')], 'i6'),
  num([t('Copie o ID do site: é o código na URL do painel ('), code('manage.wix.com/dashboard/<ID-DO-SITE>/...'), t(').')], 'i6'),
  num([t('Rode '), b('⚙️ Cadência ▸ Wix (carrinhos) ▸ Configurar API key'), t(' e cole os dois. O teste mostra quantos carrinhos estão em aberto.')], 'i6'),
  num([b('Importar carrinhos agora'), t(' coloca em cadência todos os carrinhos dos últimos 60 dias, inclusive os antigos.')], 'i6'),
  tabela(['Regra automática', 'Comportamento'], [
    ['Espera inicial', 'Carrinhos com menos de 60 min (WIX_ESPERA_MIN) aguardam o próximo ciclo'],
    ['Carrinho antigo', 'Mais de 7 dias (WIX_DIAS_CARRINHO_ANTIGO) vai para a cadência de retomada, com abordagem mais leve'],
    ['Duplicidade', 'O mesmo e-mail nunca recebe duas cadências de carrinho ao mesmo tempo'],
    ['Sem contato', 'Carrinho sem e-mail e sem telefone é ignorado'],
    ['Compra B2B', 'E-mail de domínio próprio (ex.: compras@empresa.com.br): a mensagem oferece orçamento formal e faturamento em CNPJ'],
    ['Recuperação', 'Quando o cliente conclui a compra, a cadência encerra como GANHO'],
  ], [2.6, 7.4]),
  nota('Atenção:', 'se a automação nativa de carrinho abandonado do Wix estiver ligada, desligue-a ou alinhe os textos, para o cliente não receber mensagens repetidas.'),

  h2('3.6 Ajustes finais e ativação'),
  num('Na aba Config, preencha EMAIL_REMETENTE, NOME_REMETENTE, ASSINATURA_HTML, EMAIL_RESUMO_DIARIO e os VAR_LINK_*.', 'i5'),
  num('Em VAR_LINK_AGENDA, use a página de agendamento do Google Agenda. Quando o lead agenda por ela, a cadência para sozinha.', 'i5'),
  num('Revise os textos da aba Cadencias, principalmente as orientações técnicas da linha Oxy.', 'i5'),
  num([b('Instalação ▸ 2. Testar e-mail UOL'), t(': confira se o remetente chegou correto.')], 'i5'),
  num([b('Instalação ▸ 3. Ativar automação'), t(': instala o motor horário, o resumo das 7h e o gatilho do checkbox.')], 'i5'),
);

// ── 4. Parâmetros ──
corpo.push(
  h1('4. Parâmetros da aba Config'),
  p('Todos os parâmetros podem ser alterados direto na planilha. A mudança vale a partir do próximo ciclo.'),
  tabela(['Chave', 'Valor padrão', 'Função'], CONFIG.map(([k, v, d]) => [k, String(v), d]), [3.2, 2.6, 4.6]),
);

// ── 5. Cadências ──
const grupos = {};
MODELOS.forEach(m => (grupos[m[0]] = grupos[m[0]] || []).push(m));
const titulos = {
  POS_VENDA: '5.1 Pós-venda e-commerce', PROSPECCAO_B2B: '5.2 Prospecção B2B', PROPOSTA: '5.3 Propostas comerciais',
  CARRINHO_ABANDONADO: '5.4 Carrinho abandonado (recente)', CARRINHO_RETOMADA: '5.5 Retomada de carrinho antigo',
};
corpo.push(
  h1('5. Cadências padrão'),
  p('"Dia" conta dias úteis a partir da inscrição. Os textos ficam na aba Cadencias e podem ser editados sem mexer no código. Para criar uma cadência nova, use outro nome na coluna Cadencia.'),
);
Object.keys(grupos).forEach(g => {
  corpo.push(h2(titulos[g] || g), tabela(['Passo', 'Dia', 'Canal', 'Assunto / objetivo'],
    grupos[g].map(m => [m[1], 'D+' + m[2], m[3], m[4] || m[5].substring(0, 110) + (m[5].length > 110 ? '…' : '')]),
    [0.8, 0.9, 1.5, 6.8]), esp());
});

// ── 6. Operação diária ──
corpo.push(
  h1('6. Operação diária'),
  h2('6.1 Rotina recomendada'),
  tabela(['Quando', 'Ação', 'Tempo'], [
    ['07h', 'Ler o resumo do dia que chega por e-mail', '2 min'],
    ['09h e 14h', 'Aba Fila_WhatsApp: clicar em 💬 Abrir, enviar e marcar ☑ Enviado (inclui carrinhos)', '10–15 min'],
    ['Conforme Agenda', 'Fazer as ligações agendadas (o roteiro está no evento)', 'variável'],
    ['Diário', 'Cadastrar leads B2B novos na aba Contatos', '5 min'],
    ['Semanal', 'Revisar a aba Log e os contatos com status ERRO', '10 min'],
  ], [2, 6.5, 1.5]),
  esp(),
  h2('6.2 Ações frequentes'),
  tabela(['Situação', 'Como fazer'], [
    ['Novo lead B2B', 'Aba Contatos: preencher Nome, Empresa, Email e Telefone. Deixar Status vazio (Cadencia opcional; padrão PROSPECCAO_B2B)'],
    ['Lead pediu para não ser contatado', 'Selecionar a linha ▸ ⚙️ Cadência ▸ Pausar'],
    ['Fechou negócio fora do Bling/Wix', 'Selecionar a linha ▸ ⚙️ Cadência ▸ Marcar como GANHO'],
    ['Retomar contato pausado ou com erro', 'Corrigir o dado (e-mail, telefone, cadência) ▸ selecionar ▸ Retomar'],
    ['Puxar carrinhos do Wix na hora', '⚙️ Cadência ▸ Wix (carrinhos) ▸ Importar carrinhos agora'],
    ['Forçar um ciclo agora', '⚙️ Cadência ▸ Processar agora'],
    ['Editar mensagens', 'Aba Cadencias, colunas Assunto e Mensagem'],
  ], [3.3, 6.7]),
  esp(),
  h2('6.3 Status do contato'),
  tabela(['Status', 'Significado'], [
    ['ATIVO', 'Em cadência, aguardando a data da próxima ação'],
    ['AGUARDANDO_WPP', 'WhatsApp na fila; avança quando "Enviado" é marcado'],
    ['CONCLUIDO', 'Todos os passos executados'],
    ['REUNIAO_AGENDADA', 'Contato convidado para um evento na Agenda'],
    ['RESPONDEU', 'Resposta por e-mail detectada (se ativado)'],
    ['GANHO / PERDIDO', 'Proposta aprovada ou recusada, carrinho recuperado, ou marcação manual'],
    ['CANCELADO', 'Pedido cancelado no Bling'],
    ['PAUSADO', 'Pausa manual'],
    ['ERRO', 'Problema de configuração; o motivo fica na coluna Observacoes'],
  ], [3, 7]),
  esp(),
  h2('6.4 Variáveis dos modelos'),
  tabela(['Variável', 'Conteúdo'], [
    ['{{nome}} {{primeiro_nome}} {{empresa}}', 'Dados do contato'],
    ['{{email}} {{telefone}}', 'Dados do contato'],
    ['{{numero_pedido}} {{total}} {{data_pedido}}', 'Pedido do Bling'],
    ['{{rastreio}}', 'Código de rastreio, buscado no Bling na hora do envio'],
    ['{{numero_proposta}} {{data_proposta}}', 'Proposta do Bling'],
    ['{{produtos}} {{link_checkout}} {{data_carrinho}} {{frase_b2b}}', 'Carrinho do Wix ({{total}} já vem com R$)'],
    ['{{link_avaliacao}} {{link_loja}} {{link_agenda}} {{assinatura_wpp}}', 'Vêm das chaves VAR_* da aba Config. Qualquer VAR_XXX nova vira {{xxx}}'],
  ], [4.8, 5.2]),
);

// ── 7. Limites ──
corpo.push(
  h1('7. Limites, conformidade e boas práticas'),
  tabela(['Item', 'Limite / orientação'], [
    ['E-mails por dia', 'Gmail gratuito: 100. Google Workspace: 1.500. Configurar LIMITE_EMAILS_DIA abaixo do teto'],
    ['API Bling', '3 requisições/s e 120 mil/dia. O script já espera 350 ms entre chamadas'],
    ['API Wix', 'Até 100 carrinhos por página; o script percorre as páginas automaticamente'],
    ['Tempo de execução', 'Até 6 min por ciclo. Acima de ~300 contatos ativos, reduzir BLING_JANELA_DIAS e WIX_JANELA_DIAS'],
    ['LGPD – B2B', 'Prospecção fundamentada em legítimo interesse. Atender de imediato pedidos de descadastro (Pausar)'],
    ['LGPD – clientes', 'Pós-venda e carrinho ligados à relação de compra. Recompra e avaliação devem permitir recusa'],
    ['WhatsApp', 'Evitar dezenas de envios seguidos. Personalizar a primeira mensagem para leads frios'],
    ['Entregabilidade', 'SPF/DKIM do domínio ativos no UOL Host. Evitar anexos e excesso de links'],
    ['Conteúdo técnico', 'Revisar as afirmações sobre dióxido de cloro e a linha Oxy; usar só informações de rótulo e ficha técnica'],
  ], [2.6, 7.4]),
);

// ── 8. Solução de problemas ──
corpo.push(
  h1('8. Solução de problemas'),
  tabela(['Sintoma', 'Causa provável', 'Solução'], [
    ['"não está em Enviar e-mail como"', 'Endereço UOL não cadastrado no Gmail', 'Refazer a etapa 3.2; conferir a grafia em EMAIL_REMETENTE'],
    ['"Bling não autorizado"', 'Credenciais ausentes ou refresh token vencido', 'Bling ▸ Autorizar / reconectar'],
    ['"Estado inválido" ao autorizar', 'Link de autorização antigo', 'Gerar um link novo pelo menu'],
    ['Pedidos não entram', 'IDs de situação diferentes na conta', 'Listar situações e ajustar BLING_PEDIDO_SITUACOES_ENTRADA'],
    ['Propostas não encerram', 'Nome da situação diferente', 'Conferir a coluna Situacao_Bling e ajustar BLING_PROPOSTA_*'],
    ['"Wix 401/403"', 'API key sem permissão ou ID do site errado', 'Gerar nova key com Read Orders; conferir o ID na URL do painel'],
    ['Carrinho não entra', 'Menos de 60 min, sem e-mail/telefone ou mesmo e-mail já em cadência', 'Conferir a aba Log e as chaves WIX_*'],
    ['"Telefone inválido para WhatsApp"', 'Número sem DDD ou incompleto', 'Corrigir o telefone e retomar o contato'],
    ['Nada é enviado', 'Fora do horário, fim de semana ou limite diário atingido', 'Conferir HORA_INICIO/HORA_FIM, SO_DIAS_UTEIS e a aba Log'],
    ['Checkbox não avança o contato', 'Gatilho de edição não instalado', 'Instalação ▸ 3. Ativar automação'],
  ], [3, 3.3, 3.7]),
);

// ── 9. Checklist ──
corpo.push(
  h1('9. Checklist de entrada em produção'),
  ...[
    'Planilha instalada e menu ⚙️ Cadência visível',
    'E-mail de teste recebido com o remetente UOL correto',
    'Bling conectado e situações conferidas na aba Config',
    'Wix conectado e automação nativa de carrinho do Wix revisada',
    'Textos das cinco cadências revisados (técnico e comercial)',
    'Links (avaliação, loja, agenda) e assinaturas preenchidos',
    'Teste com 1 contato interno em cada cadência (e-mail, WhatsApp, ligação)',
    'Automação ativada e resumo das 7h recebido no dia seguinte',
    'Responsável pela fila de WhatsApp definido',
  ].map(s => p([t('☐  '), t(s)], { spacing: { after: 80 } })),
);

// ── Documento ──
const lista = ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
  style: { paragraph: { indent: { left: 500, hanging: 300 } } } }] });

const doc = new Document({
  creator: 'Mayra Mucha',
  title: 'Manual – Cadência Integrada',
  styles: {
    default: { document: { run: { font: 'Calibri', size: 21, color: '1E293B' } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, color: AZUL },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: LARANJA, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 25, bold: true, color: LARANJA },
        paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: ['num', 'ciclo', 'i1', 'i2', 'i3', 'i4', 'i5', 'i6'].map(lista) },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } },
      titlePage: true },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: [t('Cadência Integrada · Manual de Instalação e Operação', { size: 16, color: '94A3B8' })] })] }),
      first: new Header({ children: [new Paragraph('')] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [t('OxHydro Saneamento · página ', { size: 16, color: '94A3B8' }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '94A3B8' })] })] }),
      first: new Footer({ children: [new Paragraph('')] }),
    },
    children: corpo,
  }],
});

Packer.toBuffer(doc).then(buf => {
  const saida = path.join(RAIZ, 'Manual_Cadencia_Integrada.docx');
  fs.writeFileSync(saida, buf);
  console.log('OK', saida, buf.length);
});
