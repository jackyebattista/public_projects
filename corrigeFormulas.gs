// =================================================================
// ÁREA DE CONFIGURAÇÃO - EDITE APENAS AS LINHAS ABAIXO
// =================================================================

/**
 * @desc O nome exato da aba (planilha) onde as fórmulas devem ser protegidas.
 * IMPORTANTE: O nome deve estar entre aspas e ser idêntico ao da sua planilha (maiúsculas e minúsculas contam).
 * @example const NOME_DA_ABA_MONITORADA = "Lançamentos";
 */
const NOME_DA_ABA_MONITORADA = "Lançamentos";


/**
 * @desc Lista das fórmulas e suas respectivas células.
 * Siga o formato: 'CÉLULA': '=SUA_FÓRMULA',
 * - A CÉLULA deve estar entre aspas simples. Ex: 'Z2'
 * - A FÓRMULA deve começar com '=' e estar entre aspas simples. Ex: '=SOMA(A:A)'
 * - Use uma vírgula (,) para separar cada item da lista, exceto o último.
 */
const FORMULAS_E_CELULAS = {

  // --- FÓRMULA 1: Busca o CENTRO DE CUSTO ---
  // OBJETIVO: Preenche o Centro de Custo com base no "Tipo de Lançamento" (coluna Q).
  // - Se Q for "Único Jurídico", busca o "Motivo" (T) na aba 'Oções forms e de-para' (colunas C:G).
  // - Se Q for "Único outra diretoria", copia o valor da coluna V.
  // - Se Q for "Rateio", apenas escreve "Rateio".
  // - PONTO DE EDIÇÃO: Se a tabela na aba 'Oções forms e de-para' mudar, ajuste o intervalo 'C:G'.
  'Z2': '=ARRAYFORMULA(IF(A2:A="";;IF(Q2:Q="Rateio"; "Rateio";IF(Q2:Q="Único Jurídico"; IFERROR(VLOOKUP(T2:T; \'Oções forms e de-para\'!C:G; 5; FALSE); "Verificar");IF(Q2:Q="Único outra diretoria"; V2:V; "")))))',


  // --- FÓRMULA 2: Busca a CONTA CONTÁBIL ---
  // OBJETIVO: Preenche a Conta Contábil.
  // - Busca o "Subgrupo" (coluna X) na aba 'Oções forms e de-para' (colunas L:M) para achar a conta.
  // - Se o lançamento for "Rateio", apenas escreve "Rateio".
  // - PONTO DE EDIÇÃO: Se a tabela na aba 'Oções forms e de-para' mudar, ajuste o intervalo 'L:M'.
  'AA2': '=ARRAYFORMULA(IF(A2:A="";;IFERROR(IF(Q2:Q="Rateio"; "Rateio";VLOOKUP(X2:X; \'Oções forms e de-para\'!L:M; 2; FALSE)); "Verificar")))',


  // --- FÓRMULA 3: Busca o APROVADOR FINANCEIRO ---
  // OBJETIVO: Identifica o Aprovador com base na Conta Contábil.
  // - Usa a Conta Contábil (coluna AA) e a procura na aba 'Validação' (colunas L:O) para achar o aprovador.
  // - PONTO DE EDIÇÃO: Se a sua matriz de aprovadores na aba 'Validação' mudar, ajuste o intervalo 'L:O'.
  'AL2': '=ARRAYFORMULA(IF(A2:A=""; "";IFERROR(VLOOKUP(AA2:AA; \'Validação\'!L:O; 3; FALSE);"Verificar")))',


  // --- FÓRMULA 4: Define o STATUS DE ARQUIVAMENTO ("Mover" ou "Manter") ---
  // OBJETIVO: Decide se a linha está pronta para ser arquivada pelo robô.
  // CONDIÇÕES: Marca como "Mover" se a DATA (B) for 4 meses ou mais antiga E o STATUS (AC) for "Finalizado" ou "cancelada".
  // - PONTO DE EDIÇÃO 1 (TEMPO): Para mudar o tempo de arquivamento, altere o número -4 (que significa 4 meses). Ex: para 6 meses, use -6.
  // - PONTO DE EDIÇÃO 2 (STATUS): Se os nomes dos status mudarem, altere os textos "Finalizado" e "cancelada" aqui na fórmula.
  'AR2': '=ARRAYFORMULA(IF(A2:A = "";;IF((B2:B <= EDATE(TODAY(); -4)) * ((AC2:AC = "Finalizado") + (AC2:AC = "cancelada"));"Mover"; "Manter")))'
};

// =================================================================
// FIM DA ÁREA DE CONFIGURAÇÃO
// Não altere o código abaixo desta linha para garantir o funcionamento.
// =================================================================


/**
 * @summary Gatilho que restaura a fórmula se for apagada/sobrescrita manualmente.
 * @param {Object} e O objeto de evento do gatilho onEdit.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() !== NOME_DA_ABA_MONITORADA) {
    return;
  }
  
  const editedCellA1 = range.getA1Notation();
  
  if (FORMULAS_E_CELULAS.hasOwnProperty(editedCellA1)) {
    
    if (range.getFormula() !== FORMULAS_E_CELULAS[editedCellA1]) {

      const formulaRow = range.getRow();
      const formulaCol = range.getColumn();
      const lastRow = sheet.getMaxRows();
      const numRowsToClear = lastRow - formulaRow;
      
      if (numRowsToClear > 0) {
        const rangeToClear = sheet.getRange(formulaRow + 1, formulaCol, numRowsToClear, 1);
        rangeToClear.clearContent();
      }
      
      range.setFormula(FORMULAS_E_CELULAS[editedCellA1]);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `A fórmula protegida da célula ${editedCellA1} foi restaurada.`, 
        'Aviso de Proteção', 
        5
      );
    }
  }
}

/**
 * @summary Restaura todas as fórmulas protegidas nas células configuradas.
 * Use esta função para garantir fórmulas após processamento em lote.
 * Ela pode ser chamada manualmente ou pelo script de automação master.
 */
function restaurarFormulasProtegidas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_DA_ABA_MONITORADA);
  if (!sheet) throw new Error('Aba monitorada não encontrada!');

  for (var celula in FORMULAS_E_CELULAS) {
    if (FORMULAS_E_CELULAS.hasOwnProperty(celula)) {
      sheet.getRange(celula).setFormula(FORMULAS_E_CELULAS[celula]);
    }
  }
}
```
