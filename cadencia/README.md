# Cadência Integrada: Google Agenda × WhatsApp Business × Bling × E-mail UOL

Motor de cadência em **Google Apps Script**. Uma **planilha Google** funciona como painel. Três fluxos vêm prontos:

| Cadência | Origem | Gatilho de entrada | Encerra quando |
|---|---|---|---|
| `POS_VENDA` | Bling: pedidos de venda | Pedido nas situações de entrada (padrão 6, 9 ou 15) | Pedido cancelado (12) ou fim dos passos |
| `PROSPECCAO_B2B` | Linha digitada na aba **Contatos** | Linha nova sem Status | Reunião na Agenda, resposta por e-mail, pausa ou fim dos passos |
| `PROPOSTA` | Bling: propostas comerciais | Proposta nova (exceto Rascunho) | Proposta Aprovada/Concluída (**GANHO**) ou Não aprovada (**PERDIDO**) |

## Como cada canal funciona

| Canal | Execução | Automação |
|---|---|---|
| **EMAIL** | Gmail envia **com remetente UOL Host** ("Enviar e-mail como") | 100 % automático |
| **WHATSAPP** | Mensagem vai para a aba **Fila_WhatsApp** com o botão 💬 *Abrir* (wa.me). Você clica, revisa, envia e marca **Enviado** | Semiautomático (sem custo e sem aprovação da Meta) |
| **LIGACAO / TAREFA** | Cria um evento na **Google Agenda** com roteiro, telefone e e-mail | 100 % automático |
| **Agenda (parada)** | Se o e-mail do contato aparece como convidado de um evento, a cadência encerra como `REUNIAO_AGENDADA` | 100 % automático |

O motor roda **a cada 1 h**. Os envios acontecem só dentro do horário comercial configurado e, por padrão, só em dias úteis. Às **7h** chega um resumo por e-mail com o funil, os WhatsApps pendentes e a agenda do dia.

---

## Instalação (cerca de 30 min)

### 1. Planilha e script
1. Crie uma planilha nova no Google Sheets, por exemplo **"Cadência OxHydro"**.
2. Abra **Extensões ▸ Apps Script**.
3. Crie os arquivos `Config.gs`, `Util.gs`, `Canais.gs`, `Bling.gs`, `Motor.gs` e `Menu.gs` e cole o conteúdo desta pasta.
4. Em **Configurações do projeto ▸ Mostrar arquivo de manifesto**, substitua o `appsscript.json` pelo desta pasta.
5. Salve, volte à planilha e recarregue. O menu **⚙️ Cadência** aparece.
6. Rode **⚙️ Cadência ▸ Instalação ▸ 1. Instalar planilha** e autorize as permissões.

> Alternativa por linha de comando: `npm i -g @google/clasp`, depois `clasp login`, `clasp create --type sheets` e `clasp push`, de dentro desta pasta.

### 2. E-mail UOL Host no Gmail
O Apps Script não abre conexão SMTP direta. Por isso o envio sai pelo Gmail, com o seu endereço UOL como remetente.

1. No Gmail, vá em **Configurações ▸ Contas e importação ▸ Enviar e-mail como ▸ Adicionar outro endereço**.
2. Informe nome e e-mail UOL (ex.: `contato@oxhydro.com.br`) e **desmarque** "Tratar como alias".
3. Preencha o SMTP:

| Campo | Valor |
|---|---|
| Servidor SMTP | `smtps.uhserver.com` |
| Porta | `465` com **SSL** (alternativa: `smtp.uhserver.com`, porta `587`, TLS) |
| Usuário | e-mail completo |
| Senha | senha da caixa UOL |

4. Confirme pelo código que chega na caixa UOL.
5. **Opcional, para detectar respostas:** em *Verificar o e-mail de outras contas*, adicione POP3 `pop.uhserver.com`, porta `995`, SSL. Depois coloque `DETECTAR_RESPOSTA_EMAIL = SIM` na aba Config.

> **SPF:** para não cair em spam, o SPF do domínio no painel UOL Host precisa autorizar o servidor UOL. Como o envio é feito pelo SMTP do UOL, o padrão do UOL Host já atende.

### 3. App da Web (callback do Bling)
1. No Apps Script, vá em **Implantar ▸ Nova implantação ▸ Tipo: App da Web**.
2. Escolha *Executar como: Eu* e *Quem pode acessar: Somente eu*.
3. Copie a URL que termina em **`/exec`**.

### 4. Aplicativo no Bling
1. No Bling, vá em **Central de Extensões ▸ Área do Integrador ▸ Criar aplicativo** (tipo *Aplicativo próprio*).
2. Em **Link de redirecionamento**, cole a URL `/exec`.
3. Escopos de **leitura**: *Pedidos de Venda*, *Propostas Comerciais*, *Contatos*, *Situações* e *Dados básicos da empresa*.
4. Copie o **Client ID** e o **Client Secret**.
5. Na planilha:
   1. **⚙️ Cadência ▸ Bling ▸ Configurar credenciais**: cole o ID e o Secret, depois clique em **Autorizar no Bling**.
   2. **Bling ▸ Testar conexão**.
   3. **Bling ▸ Listar situações (IDs)**: confira, na aba `Situacoes_Bling`, os IDs da sua conta e ajuste `BLING_PEDIDO_SITUACOES_*` na aba **Config**.

As credenciais e os tokens ficam em *Script Properties*, nunca na planilha. O token é renovado sozinho. Se o Bling ficar mais de 30 dias sem sincronizar, rode **Autorizar / reconectar**.

### 5. Ajustes finais
1. Na aba **Config**, preencha `EMAIL_REMETENTE`, `ASSINATURA_HTML`, `EMAIL_RESUMO_DIARIO` e os `VAR_LINK_*`. Para `VAR_LINK_AGENDA`, use a **página de agendamento** do Google Agenda: quando o lead agenda por ela, a cadência para sozinha.
2. **Instalação ▸ 2. Testar e-mail UOL**.
3. **Instalação ▸ 3. Ativar automação (gatilhos)**.

---

## Uso no dia a dia

| Ação | Como fazer |
|---|---|
| Novo lead B2B | Na aba **Contatos**, preencha Nome, Empresa, Email e Telefone. A coluna Cadencia é opcional (o padrão é `PROSPECCAO_B2B`). Deixe **Status vazio** |
| Enviar os WhatsApps | Na aba **Fila_WhatsApp**: 💬 *Abrir* ▸ enviar ▸ marcar ☑ **Enviado**. O contato avança para o próximo passo |
| Pausar, retomar ou marcar ganho | Selecione as linhas em Contatos ▸ menu **⚙️ Cadência** |
| Editar mensagens | Aba **Cadencias**. `Dia` = dias (úteis) contados a partir da inscrição |
| Criar cadência nova | Na aba Cadencias, use um nome novo na coluna `Cadencia` e informe esse nome na coluna Cadencia do contato |
| Auditoria | Aba **Log** (cada envio, erro e encerramento) |

### Variáveis disponíveis nos modelos
`{{nome}}` `{{primeiro_nome}}` `{{empresa}}` `{{email}}` `{{telefone}}`
Pedidos: `{{numero_pedido}}` `{{total}}` `{{data_pedido}}` `{{rastreio}}` (buscado no Bling na hora do envio).
Propostas: `{{numero_proposta}}` `{{total}}` `{{data_proposta}}`.
Toda chave `VAR_XXX` da aba Config vira `{{xxx}}`. Exemplos: `{{link_avaliacao}}`, `{{link_loja}}`, `{{link_agenda}}`, `{{assinatura_wpp}}`.

### Status do contato
`ATIVO` → `AGUARDANDO_WPP` (esperando você marcar Enviado) → … → `CONCLUIDO`.
Paradas: `REUNIAO_AGENDADA`, `RESPONDEU`, `GANHO`, `PERDIDO`, `CANCELADO`, `PAUSADO`, `ERRO` (o motivo fica em *Observacoes*).

---

## Limites e cuidados

| Item | Limite |
|---|---|
| E-mails por dia | Gmail gratuito: **100** · Google Workspace: **1.500**. Configure em `LIMITE_EMAILS_DIA` |
| Bling | 3 req/s e 120 mil req/dia. O script espera 350 ms entre chamadas |
| Execução do Apps Script | 6 min por ciclo. Em volumes altos (>300 contatos ativos), reduza `BLING_JANELA_DIAS` |
| LGPD | Prospecção B2B fundamentada em legítimo interesse. Respeite pedidos de descadastro: marque o contato como `PAUSADO` |
| WhatsApp | Abrir o link no computador usa o WhatsApp Business Desktop/Web. Evite disparar dezenas de mensagens em sequência para não ser marcado como spam |

## Estrutura

```
cadencia/
├── appsscript.json  manifesto (fuso São Paulo, escopos, App da Web)
├── Config.gs        abas, parâmetros padrão e cadências-modelo
├── Util.gs          planilha, datas/dias úteis, telefone, modelos, log
├── Canais.gs        e-mail UOL, fila WhatsApp, Google Agenda
├── Bling.gs         OAuth 2.0, API v3, sincronização de pedidos e propostas
├── Motor.gs         motor da cadência, gatilho de edição, resumo diário
└── Menu.gs          menu, gatilhos, ações em contatos selecionados
```
