
/**SCRIPTS CRIADOS PARA A PLANILHA CONTROLE PROCESSUAL 2025*/
/*************
* UI & Events*
*************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  calculateTotal();

  /**
   * Adiciona menus e itens que chamam as funções do script
   * os menus ficam visíveis em todas as abas e tem itens
   * nos itens é que são atribuídas as funções
   */
  ui
    .createMenu("Recebimentos")
    .addItem("Adicionar linha processos recebidos", "addLineReceivedProcessesSheet")
    .addItem("Deletar linha processos recebidos", "deleteLastLineReceivedProcessesSheet")
    .addToUi();

  ui
    .createMenu("Criados")
    .addItem("Adicionar linha processos criados", "addLineCreatedProcessesSheet")
    .addItem("Deletar linha processos criados", "deleteLastLineCreatedProcessesSheet")
    .addToUi();

  ui
    .createMenu("Adicionar Linha Plano de Trabalho")
    .addItem("Adicionar Linha Janeiro", "handleAddJan")
    .addItem("Adicionar Linha Fevereiro", "handleAddFeb")
    .addItem("Adicionar Linha Março", "handleAddMar")
    .addItem("Adicionar Linha Abril", "handleAddApr")
    .addItem("Adicionar Linha Maio", "handleAddMay")
    .addItem("Adicionar Linha Junho", "handleAddJun")
    .addItem("Adicionar Linha Julho", "handleAddJul")
    .addItem("Adicionar Linha Agosto", "handleAddAug")
    .addItem("Adicionar Linha Setembro", "handleAddSep")
    .addItem("Adicionar Linha Outubro", "handleAddOct")
    .addItem("Adicionar Linha Novembro", "handleAddNov")
    .addItem("Adicionar Linha Dezembro", "handleAddDec")
    .addToUi();

  ui
    .createMenu("Remover Linha Plano de Trabalho")
    .addItem("Remover Linha Janeiro", "handleRemoveLastLineJan")
    .addItem("Remover Linha Fevereiro", "handleRemoveLastLineFeb")
    .addItem("Remover Linha Março", "handleRemoveLastLineMar")
    .addItem("Remover Linha Abril", "handleRemoveLastLineApr")
    .addItem("Remover Linha Maio", "handleRemoveLastLineMay")
    .addItem("Remover Linha Junho", "handleRemoveLastLineJun")
    .addItem("Remover Linha Julho", "handleRemoveLastLineJul")
    .addItem("Remover Linha Agosto", "handleRemoveLastLineAug")
    .addItem("Remover Linha Setembro", "handleRemoveLastLineSep")
    .addItem("Remover Linha Outubro", "handleRemoveLastLineOct")
    .addItem("Remover Linha Novembro", "handleRemoveLastLineNov")
    .addItem("Remover Linha Dezembro", "handleRemoveLastLineDec")
    .addToUi();
}
/*************
* UI & Events*
*************/

/********************
* Plano de Trabalho *
********************/
function addLineWorkPlan(month) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANO DE TRABALHO (2025)");
  let lastRowMonth = findLastRowMonth(month);

  /**Cria o mês, caso não exista*/
  if (lastRowMonth == -1) {
    let newRowIndex = sheet.getLastRow() + 1;
    sheet.insertRowAfter(sheet.getLastRow());
    sheet.getRange(newRowIndex, 1).setValue(month);
    sheet.getRange(newRowIndex, 6).setValue(month.toUpperCase());
    sheet.getRange(newRowIndex, 1, 1, sheet.getLastColumn()).setBackground(decideColor(month));
  } else {
    let newRowIndex = lastRowMonth + 1;

    sheet.insertRowAfter(lastRowMonth);

    let upperRow = sheet.getRange(lastRowMonth, 1, 1, sheet.getLastColumn());
    let newRow = sheet.getRange(newRowIndex, 1, 1, sheet.getLastColumn());
    /**Insere o nome do mês na coluna F, já que não é só na primeira coluna que tem o nome do mês*/ 
    sheet.getRange(newRow.getRow(), 6).setValue(month.toUpperCase());

    /**Copiar formatação da linha de cima
     *Parâmetros da função copyFormatToRange: id da planilha, primeira e última colunas do intervalo
     *a ser copiado respectivamente, primeira e última linhas do intervalo a ser copiado
     *respectivamente
    */
    upperRow.copyFormatToRange(sheet, 1, sheet.getLastColumn(), newRowIndex, newRowIndex);

    /**Seleciona a região mesclada com base no nome do mês
     * para inserir a nova linha do mês no final
    */
    let mergedRanges = sheet.getRange(lastRowMonth, 1).getMergedRanges();

    if (mergedRanges.length > 0) {
      let mergedRange = mergedRanges[0];

      /**Seleciona a primeira linha da mescla encontrada,
       * adiciona +1 no número de linhas da mescla
      */
      let firstRow = mergedRange.getRow();
      let numRows = mergedRange.getNumRows() + 1;

      /**Seleciona um intervalo que é a soma da mescla anterior
       * com a nova linha e mescla esse novo intervalo
       */
      let newMergeRange = sheet.getRange(firstRow, 1, numRows);
      newMergeRange.merge();
    } else {
       /**Caso não haja mesclagem anterior, mescla a linha original com a nova linha
       * Aqui mescla a linha anterior à nova linha,
       * caso um mês já exista mas esteja sem nenhuma mesclagem
       */
      let mergeRange = sheet.getRange(lastRowMonth, 1, 2);
      mergeRange.merge();
    }
  }
}

function deleteLastLineWorkPlan(month) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANO DE TRABALHO (2025)");
  let lastRowMonth = findLastRowMonth(month);

  if (lastRowMonth == -1) {
    return;
  }

  sheet.deleteRow(lastRowMonth);

  return;
}
/********************
* Plano de Trabalho *
********************/

/****************************************
* Plano de Trabalho: Funções auxiliares *
****************************************/
/**Encontra a primeira linha com o nome de um mês*/
function findFirstRowMonth(data, month) {
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === month) {
      return i + 1;
    }
  }

  return -1;
}

/**Encontra a última linha com o nome de um mês*/
function findLastRowMonth(month) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANO DE TRABALHO (2025)");
  let range = sheet.getDataRange();
  let mergedRanges = range.getMergedRanges();

  for (let i = 0; i < mergedRanges.length; i++) {
    let mergedRange = mergedRanges[i];

    /**Verifica se a célula inicial do intervalo mesclado é o nome mês */
    if (mergedRange.getCell(1, 1).getValue() === month) {
      /**Retorna a última linha da mesclagem, a última aparição de um mês */
      return mergedRange.getLastRow();
    }
  }

  /**Pegar linha individual se for a última linha restante de um mês */
  for(let i = 0; i < range.getLastRow(); i++) {
    let cellWithMonthName = sheet.getRange(i + 1, 1).getValue();

    if(cellWithMonthName === month) {
      return i + 1;
    }
  }

  return -1; /**Retorna -1 se não encontrar*/
}

function calculateTotal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANO DE TRABALHO (2025)");
  const data = sheet.getDataRange().getValues();

  let month = "";
  let mediumByMonths = {};

  for (let i = 2; i < data.length; i++) {
    /**Atribui o nome do mês à month */
    if (data[i][0].length > 0) {
      month = data[i][0];
    }

    /**Pega o valor da coluna Estimativa na linha i*/
    const value = data[i][8];
    if (value !== null && value > 0) {
      /**mediumByMonths é um objeto, que contém vários meses {Janeiro: 1000, Fevereiro: 500} 
       * o if e else abaixo serve para criar a chave MÊS, caso ela não exista e atribuir o valor
       * de Estimativa ou acumular esse valor, caso a chave MÊS já exista
      */
      if (mediumByMonths[month]) {
        mediumByMonths[month] = mediumByMonths[month] + value;
      } else {
        mediumByMonths[month] = value;
      }
    }
  }

  /**Título dos resumos (Totais) */
  sheet.getRange(2, 12).setValue("Meses").setBackground("#808080").setFontColor("#fff");
  sheet.getRange(2, 13).setValue("Valor Total").setBackground("#808080").setFontColor("#fff");
  sheet.getRange(2, 15).setValue("Total").setBackground("#808080").setFontColor("#fff");

  let totalMedium = 0;

  /**Calcula a soma da média total de todos os meses*/
  for (const key in mediumByMonths) {
    if (!Number.isNaN(mediumByMonths[key])) {
      totalMedium = totalMedium + mediumByMonths[key];
    }
  }

  sheet.getRange(3, 15).setValue(totalMedium).setNumberFormat('R$ 0,000.00').setBackground("#808080").setFontColor("#fff");

  /**Essa parte monta o resumo total por mês */
  let interval = 14;
  let aux = 0;
  for (let i = 2; i <= interval; i++) {
    switch (aux) {
      case 1:
        sheet.getRange(i, 12).setValue("Janeiro").setBackground(decideColor("Janeiro"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Janeiro"] ? mediumByMonths["Janeiro"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Janeiro"));
        break;
      case 2:
        sheet.getRange(i, 12).setValue("Fevereiro").setBackground(decideColor("Fevereiro"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Fevereiro"] ? mediumByMonths["Fevereiro"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Fevereiro"));
        break;
      case 3:
        sheet.getRange(i, 12).setValue("Março").setBackground(decideColor("Março"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Março"] ? mediumByMonths["Março"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Março"));
        break;
      case 4:
        sheet.getRange(i, 12).setValue("Abril").setBackground(decideColor("Abril"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Abril"] ? mediumByMonths["Abril"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Abril"));
        break;
      case 5:
        sheet.getRange(i, 12).setValue("Maio").setBackground(decideColor("Maio"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Maio"] ? mediumByMonths["Maio"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Maio"));
        break;
      case 6:
        sheet.getRange(i, 12).setValue("Junho").setBackground(decideColor("Junho"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Junho"] ? mediumByMonths["Junho"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Junho"));
        break;
      case 7:
        sheet.getRange(i, 12).setValue("Julho").setBackground(decideColor("Julho"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Julho"] ? mediumByMonths["Julho"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Julho"));
        break;
      case 8:
        sheet.getRange(i, 12).setValue("Agosto").setBackground(decideColor("Agosto"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Agosto"] ? mediumByMonths["Agosto"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Agosto"));
        break;
      case 9:
        sheet.getRange(i, 12).setValue("Setembro").setBackground(decideColor("Setembro"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Setembro"] ? mediumByMonths["Setembro"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Setembro"));
        break;
      case 10:
        sheet.getRange(i, 12).setValue("Outubro").setBackground(decideColor("Outubro"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Outubro"] ? mediumByMonths["Outubro"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Outubro"));
        break;
      case 11:
        sheet.getRange(i, 12).setValue("Novembro").setBackground(decideColor("Novembro"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Novembro"] ? mediumByMonths["Novembro"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Novembro"));
        break;
      case 12:
        sheet.getRange(i, 12).setValue("Dezembro").setBackground(decideColor("Dezembro"));
        sheet
          .getRange(i, 13)
          .setValue(mediumByMonths["Dezembro"] ? mediumByMonths["Dezembro"] : 0)
          .setNumberFormat('R$ 0,000.00')
          .setBackground(decideColor("Dezembro"));
        break;
    }
    aux++;
  }
}
/****************************************
* Plano de Trabalho: Funções auxiliares *
****************************************/

/***********************
* Planejamento Semanal *
***********************/
function addNewRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANEJAMENTO SEMANAL (2025)");

  const lastRow = sheet.getLastRow();

  const rangeLastRow = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  /**newRange é a localização da nova linha a ser adicionada */
  const newRange = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());

  /**Copia a formatação da última linha e cola na nova linha */
  rangeLastRow.copyTo(newRange, { formatOnly: true });

  sheet.insertRowAfter(lastRow);

  const description = "Descrição";
  sheet.getRange(lastRow + 1, 1).setValue(description);
}

function removeLastRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANEJAMENTO SEMANAL (2025)");

  const lastRow = sheet.getLastRow();

  /**Deleta a última lnha caso ela não seja a primeira linha após o cabeçalho */
  if (lastRow > 3) {
    sheet.deleteRow(lastRow);
  }
}
/***********************
* Planejamento Semanal *
***********************/

/*************************
* Recebimentos e Criados *
*************************/
/**Adiciona linha na primeira aba*/
function addLineReceivedProcessesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RECEBIMENTOS (2025)");

  const lastRow = sheet.getLastRow();

  const rangeLastRow = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  /**newRange é a localização da nova linha a ser adicionada */
  const newRange = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());

  /**Copia a formatação da última linha e cola na nova linha */
  rangeLastRow.copyTo(newRange, { formatOnly: true });

  sheet.insertRowAfter(lastRow);

  /**Ao inserir uma nova linha, a data de hoje é adicionada à coluna 1, na última linha,
   * para que esta não fique vazia, e ao excluir, não gere erro
   */
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  sheet.getRange(lastRow + 1, 1).setValue(currentDate);
}

/**Remove última linha na primeira aba*/
function deleteLastLineReceivedProcessesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RECEBIMENTOS (2025)");

  const lastRow = sheet.getLastRow();

  /**Deleta a última lnha caso ela não seja a primeira linha após o cabeçalho */
  if (lastRow > 3) {
    sheet.deleteRow(lastRow);
  }
}

/**Adiciona linha na segunda aba*/
function addLineCreatedProcessesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRIADOS (2025)");

  const lastRow = sheet.getLastRow();

  const rangeLastRow = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
   /**newRange é a localização da nova linha a ser adicionada */
  const newRange = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());

  /**Copia a formatação da última linha e cola na nova linha */
  rangeLastRow.copyTo(newRange, { formatOnly: true });
  sheet.insertRowAfter(lastRow);

  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  sheet.getRange(lastRow + 1, 1).setValue(currentDate);
}

/**Remove última linha na segunda aba*/
function deleteLastLineCreatedProcessesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRIADOS (2025)");

  const lastRow = sheet.getLastRow();

  /**Deleta a última lnha caso ela não seja a primeira linha após o cabeçalho */
  if (lastRow > 3) {
    sheet.deleteRow(lastRow);
  }
}
/*************************
* Recebimentos e Criados *
*************************/

/********
* Utils *
********/
/**Função que decide a cor baseado no nome do mês que é passado na chamada da função
 * retorna uma string hexadecimal que equivale a um tom de cor
 */
function decideColor(month) {
  switch (month) {
    case "Janeiro":
      return "#d9ead3";

    case "Fevereiro":
      return "#c9daf8";

    case "Março":
      return "#fce5cd";

    case "Abril":
      return "#f4cccc";

    case "Maio":
      return "#eeeeee";

    case "Junho":
      return "#93c47d";

    case "Julho":
      return "#5EEAD4";

    case "Agosto":
      return "#FACC15";

    case "Setembro":
      return "#F0ABFC";

    case "Outubro":
      return "#CBD5E1";

    case "Novembro":
      return "#BEF264";

    case "Dezembro":
      return "#F43F5E";
  }
}

/**As funções abaixo simplesmente chamam as funções
 * addLineWorkPlan e deleteLastLineWorkPlan passando
 * um mês na chamada da função
*/
function handleAddJan() {
  addLineWorkPlan("Janeiro");
}

function handleAddFeb() {
  addLineWorkPlan("Fevereiro");
}

function handleAddMar() {
  addLineWorkPlan("Março");
}

function handleAddApr() {
  addLineWorkPlan("Abril");
}

function handleAddMay() {
  addLineWorkPlan("Maio");
}

function handleAddJun() {
  addLineWorkPlan("Junho");
}

function handleAddJul() {
  addLineWorkPlan("Julho");
}

function handleAddAug() {
  addLineWorkPlan("Agosto");
}

function handleAddSep() {
  addLineWorkPlan("Setembro");
}

function handleAddOct() {
  addLineWorkPlan("Outubro");
}

function handleAddNov() {
  addLineWorkPlan("Novembro");
}

function handleAddDec() {
  addLineWorkPlan("Dezembro");
}

function handleRemoveLastLineJan() {
  deleteLastLineWorkPlan("Janeiro");
}

function handleRemoveLastLineFeb() {
  deleteLastLineWorkPlan("Fevereiro");
}

function handleRemoveLastLineMar() {
  deleteLastLineWorkPlan("Março");
}

function handleRemoveLastLineApr() {
  deleteLastLineWorkPlan("Abril");
}

function handleRemoveLastLineMay() {
  deleteLastLineWorkPlan("Maio");
}

function handleRemoveLastLineJun() {
  deleteLastLineWorkPlan("Junho");
}

function handleRemoveLastLineJul() {
  deleteLastLineWorkPlan("Julho");
}

function handleRemoveLastLineAug() {
  deleteLastLineWorkPlan("Agosto");
}

function handleRemoveLastLineSep() {
  deleteLastLineWorkPlan("Setembro");
}

function handleRemoveLastLineOct() {
  deleteLastLineWorkPlan("Outubro");
}

function handleRemoveLastLineNov() {
  deleteLastLineWorkPlan("Novembro");
}

function handleRemoveLastLineDec() {
  deleteLastLineWorkPlan("Dezembro");
}
/********
* Utils *
********/
