# autofill-forms
Código usado no Apps Scripts para Auto preencher o Google Forms com base em dados no Google Sheets


```
const LINK_PLANILHA = 'link planilha';
const LINK_FORM = 'link do forms';
const ABA_DADOS = 'Dados';
const ABA_MAPEAMENTO = 'Link'; // Aba onde os índices são configurados
const COLUNA_MAPEAMENTO = 1; // Número da coluna onde estão os índices das colunas de dados
const COLUNA_INDICE_FORMS = 2; // Número da coluna onde estão os índices do Forms

function atualizarFormulario() {
    const ss = SpreadsheetApp.openByUrl(LINK_PLANILHA);
    const form = FormApp.openByUrl(LINK_FORM);
    const sheetData = ss.getSheetByName(ABA_DADOS);
    const sheetMap = ss.getSheetByName(ABA_MAPEAMENTO);

    if (!sheetData || !sheetMap) {
      Logger.log("Aba de dados ou de mapeamento não encontrada.");
      return;
    }

    // Obter mapeamento das colunas da planilha e os índices do Forms
    const mapeamento = sheetMap.getRange(2, COLUNA_MAPEAMENTO, sheetMap.getLastRow() - 1, 2).getValues();
    
    for (let i = 0; i < mapeamento.length; i++) {
      try {
        const colunaDados = parseInt(mapeamento[i][0]);
        const indiceForms = parseInt(mapeamento[i][1]);

        if (isNaN(colunaDados) || isNaN(indiceForms)) {
          Logger.log(`Linha ${i + 2}: Índice inválido no mapeamento. Pulando...`);
          continue;
        }

        const listForm = form.getItems();
        if (indiceForms >= listForm.length) {
          Logger.log(`Índice ${indiceForms}: Pergunta não encontrada no Forms.`);
          continue;
        }

        const lastRow = getUltimaLinhaPreenchida(sheetData, colunaDados);
        if (lastRow < 2) {
          let listItem = listForm[indiceForms];
          if (listItem.getType() === FormApp.ItemType.LIST) {
            listItem.asListItem().setChoiceValues(['']);
            Logger.log(`Coluna ${colunaDados} → Forms Índice ${indiceForms} (Lista)`);
          } else if (listItem.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
            listItem.asMultipleChoiceItem().setChoiceValues(['']);
            Logger.log(`Coluna ${colunaDados} → Forms Índice ${indiceForms} (Múltipla Escolha)`);
          } else {
            Logger.log(`Índice ${indiceForms}: Tipo de pergunta incompatível.`);
          }
        }

        let valores = sheetData.getRange(2, colunaDados, lastRow - 1, 1).getValues();
        
        // Remover valores inválidos e vazios
        valores = valores.flat().map(val => val.toString().trim()).filter(val => val !== '');

        if (valores.length === 0) {
          Logger.log(`Não há dados`);
          continue;
        }

        const listItem = listForm[indiceForms];
        if (listItem.getType() === FormApp.ItemType.LIST) {
          listItem.asListItem().setChoiceValues(valores);
          Logger.log(`Coluna ${colunaDados} → Forms Índice ${indiceForms} (Lista)`);
        } else if (listItem.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          listItem.asMultipleChoiceItem().setChoiceValues(valores);
          Logger.log(`Coluna ${colunaDados} → Forms Índice ${indiceForms} (Múltipla Escolha)`);
        } else {
          Logger.log(`Índice ${indiceForms}: Tipo de pergunta incompatível.`);
        }
      } catch (e) {
        Logger.log("Erro ao atualizar formulário: " + e.message);
        }
      }
}

function getUltimaLinhaPreenchida(sheet, coluna) {
  const valores = sheet.getRange(1, coluna, sheet.getMaxRows(), 1).getValues();
  return valores.flat().map(val => val.toString().trim()).filter(val => val !== '').length;
}


```
