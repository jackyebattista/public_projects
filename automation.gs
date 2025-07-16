// ========================================================================
// CONFIGURAÇÃO PRINCIPAL - EDITE APENAS AQUI
// ========================================================================

const URL_PLANILHA_ORIGEM  = 'LINK DA PLANILHA DE ORIGEM';
const URL_PLANILHA_DESTINO = 'LINK DA PLANILHA DE DESTINO';

const NOME_ABA_ORIGEM      = "NOME DA ABA";
const NOME_ABA_DESTINO     = "NOME DA ABA";
const COLUNA_GATILHO       = "AR"; // Coluna que contém o texto para mover
const TEXTO_PARA_MOVER     = "Mover";

const LISTA_EMAILS_NOTIFICACAO = "raquel_indelicato@whirlpool.com,esther_reis@whirlpool.com,guilherme_vaz@whirlpool.com,milena_borges@whirlpool.com,eduardo_ribeiro@whirlpool.com,patricia_c_mattos@whirlpool.com,jaqueline_batista_jbplan@whirlpool.com";

// ========================================================================
// FUNÇÕES DE INTERFACE DO USUÁRIO (UI)
// ========================================================================

/**
 * Cria o menu personalizado ao abrir a planilha.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Histórico')
    .addItem('▶️ Registrar Histórico', 'showSidebar')
    .addSeparator()
    .addItem('⚙️ Criar Gatilho Mensal', 'criarTriggerMensal')
    .addToUi();
}

/**
 * Mostra a barra lateral com a interface do arquivo index.html.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Registrar Histórico');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ========================================================================
// PONTOS DE ENTRADA DA AUTOMAÇÃO
// ========================================================================

/**
 * Função chamada pelo gatilho automático mensal.
 */
function processamentoMensal() {
  executarProcessoHistorico('automatico');
}

/**
 * Função chamada pelo botão na barra lateral (sidebar).
 */
function comExecHistoricoManual() {
  return executarProcessoHistorico('manual');
}

// ========================================================================
// ORQUESTRADOR PRINCIPAL DO PROCESSO
// ========================================================================

/**
 * Orquestra todo o processo de arquivamento, com status visuais.
 * @param {string} modo - Indica se a execução é 'manual' ou 'automatico'.
 */
function executarProcessoHistorico(modo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Iniciando processo...', 'Arquivamento', 5);

    const ssOrigem = SpreadsheetApp.openByUrl(URL_PLANILHA_ORIGEM);
    const abaOrigem = ssOrigem.getSheetByName(NOME_ABA_ORIGEM);
    if (!abaOrigem) throw new Error(`Aba "${NOME_ABA_ORIGEM}" não encontrada!`);

    const dataProcessamento = getDataAtual();

    ss.toast('Lendo e separando dados...', 'Arquivamento', 5);
    const resultado = separarDados(abaOrigem, dataProcessamento);
    
    if (resultado.status === 'nao_ha_dados') {
      ss.toast('Nenhum dado para mover.', 'Arquivamento', 5);
      return {status: 'nao_ha_dados'};
    }
    
    ss.toast(`Copiando ${resultado.linhasParaMover.length} linha(s) para o destino...`, 'Arquivamento', 10);
    copiarParaDestino(resultado.linhasParaMover);

    ss.toast('Limpando e reescrevendo a planilha de origem...', 'Arquivamento', 10);
    reescreverOrigem(abaOrigem, resultado.linhasParaManter);

    ss.toast('Restaurando fórmulas...', 'Arquivamento', 5);
    
    // --- AJUSTE PRINCIPAL ---
    // A chamada da função agora corresponde exatamente à função disponível no arquivo 'corrigeFormula.gs'.
    if (typeof restaurarFormulasProtegidas === "function") {
      restaurarFormulasProtegidas();
    } else {
      Logger.log("AVISO: Função 'restaurarFormulasProtegidas' não encontrada. Verifique se o arquivo 'corrigeFormula.gs' está no projeto.");
    }

    ss.toast('Enviando e-mail de confirmação...', 'Arquivamento', 5);
    enviarEmailConfirmacao(resultado.linhasParaMover, modo, dataProcessamento);

    ss.toast('Processo concluído com sucesso!', 'Arquivamento', 5);
    return {status: 'ok'};

  } catch (e) {
    Logger.log(e.stack); // Loga os detalhes do erro para depuração.
    ss.getUi().alert(`Ocorreu um erro: ${e.message}`);
    return {status: 'erro', message: e.message};
  }
}

// ========================================================================
// FUNÇÕES DE PROCESSAMENTO DE DADOS
// ========================================================================

/**
 * Lê os dados da aba de origem e separa em dois arrays.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} aba A aba de origem.
 * @param {string} dataAtual A data/hora do processamento para adicionar aos dados movidos.
 * @returns {object} Um objeto contendo o status e os arrays de dados.
 */
function separarDados(aba, dataAtual) {
  const lastRow = aba.getLastRow();
  if (lastRow < 2) return { status: 'nao_ha_dados' };

  const firstCol = 1;
  const lastCol = colunaLetraParaNumero(COLUNA_GATILHO);

  const original = aba.getRange(2, firstCol, lastRow - 1, lastCol).getValues();
  const dadosColunaMover = aba.getRange(2, lastCol, lastRow - 1, 1).getValues();

  const linhasParaMover = [];
  const linhasParaManter = [];

  for (let i = 0; i < original.length; i++) {
    if (dadosColunaMover[i][0].toString().trim() === TEXTO_PARA_MOVER) {
      linhasParaMover.push(original[i].concat([dataAtual]));
    } else {
      linhasParaManter.push(original[i]);
    }
  }

  if (linhasParaMover.length === 0) {
    return { status: 'nao_ha_dados' };
  }

  return { status: 'ok', linhasParaMover, linhasParaManter };
}

/**
 * Escreve os dados movidos para a planilha de destino.
 */
function copiarParaDestino(linhasParaMover) {
  const ssDestino = SpreadsheetApp.openByUrl(URL_PLANILHA_DESTINO);
  const abaDestino = ssDestino.getSheetByName(NOME_ABA_DESTINO);
  if (!abaDestino) throw new Error(`Aba "${NOME_ABA_DESTINO}" não encontrada na planilha de destino!`);
  
  abaDestino.getRange(
    abaDestino.getLastRow() + 1, 1, 
    linhasParaMover.length, 
    linhasParaMover[0].length
  ).setValues(linhasParaMover);
}

/**
 * Limpa a aba de origem e reescreve os dados que não foram movidos.
 */
function reescreverOrigem(aba, linhasParaManter) {
  const ultimaLinhaAtual = aba.getLastRow();
  if (ultimaLinhaAtual > 1) {
    aba.getRange(2, 1, ultimaLinhaAtual - 1, aba.getLastColumn()).clearContent();
  }

  if (linhasParaManter.length > 0) {
    aba.getRange(2, 1, linhasParaManter.length, linhasParaManter[0].length).setValues(linhasParaManter);
  }
}

// ========================================================================
// FUNÇÕES AUXILIARES E DE NOTIFICAÇÃO
// ========================================================================

/**
 * Envia um e-mail de confirmação com um anexo CSV.
 */
function enviarEmailConfirmacao(linhasMovidasNoProcesso, modo, dataAtual) {
  if (!linhasMovidasNoProcesso || linhasMovidasNoProcesso.length === 0) return;

  const csv = linhasMovidasNoProcesso.map(row => row.join(',')).join('\n');
  const anexo = Utilities.newBlob(csv, "text/csv", "Dados_Arquivados_" + dataAtual.replace(/[\/ :]/g, "_") + ".csv");
  
  const tipoProc = (modo === 'automatico') ? "PROCESSO AUTOMÁTICO (agendado)" : "PROCESSO MANUAL (botão)";
  const assunto = `${tipoProc} | Planilha CSV Formularios notas RPA ${dataAtual}`;
  
  const corpo = `Este e-mail foi gerado por: ${tipoProc}\n\nSegue em anexo a planilha CSV com os dados da planilha Formularios notas RPA que foram arquivados na data de: ${dataAtual}`;

  MailApp.sendEmail({
    to: LISTA_EMAILS_NOTIFICACAO,
    subject: assunto,
    body: corpo,
    attachments: [anexo]
  });
}

/**
 * Converte a letra da coluna para o número correspondente.
 */
function colunaLetraParaNumero(coluna) {
  let numero = 0;
  for (let i = 0; i < coluna.length; i++) {
    numero = numero * 26 + (coluna.charCodeAt(i) - 64);
  }
  return numero;
}

/**
 * Retorna a data e hora atuais formatadas como uma string.
 */
function getDataAtual() {
  const dataAtual = new Date();
  const dia = dataAtual.getDate().toString().padStart(2, '0');
  const mes = (dataAtual.getMonth() + 1).toString().padStart(2, '0');
  const ano = dataAtual.getFullYear();
  const hora = dataAtual.getHours().toString().padStart(2, '0');
  const minuto = dataAtual.getMinutes().toString().padStart(2, '0');
  return `${dia}/${mes}/${ano} ${hora}:${minuto}`;
}

// ========================================================================
// FUNÇÃO DE CONFIGURAÇÃO DE GATILHO (TRIGGER)
// ========================================================================

/**
 * Cria um gatilho mensal para executar o processo automaticamente.
 */
function criarTriggerMensal() {
  const functionName = 'processamentoMensal';
  
  // Remove gatilhos antigos para evitar duplicidade
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Cria um novo gatilho para rodar todo dia 1 de cada mês
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onMonthDay(1)
    .atHour(7)
    .create();
    
  SpreadsheetApp.getUi().alert(`Gatilho mensal criado com sucesso para a função "${functionName}"!`);
}
