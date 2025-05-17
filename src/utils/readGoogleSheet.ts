import {
  GoogleSpreadsheet,
  GoogleSpreadsheetWorksheet,
} from "google-spreadsheet";
import { google } from "googleapis";

interface ComparisonOptions {
  compareOrder?: boolean; // Indica se a ordem dos valores deve ser levada em consideração
}

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

export async function compareColumnWithArray(
  spreadsheetId: string,
  columnHeader: string,
  expectedValues: any[],
  sheetNameOrIndex: string | number = 0, // Permite especificar aba por nome ou índice (default: 0)
): Promise<null | Record<string, string[]>> {
  try {
    //const client = await auth.getClient();

    // Inicializa o documento do Google Sheets com o ID
    const doc = new GoogleSpreadsheet(spreadsheetId, auth);

    // Carrega as informações da planilha
    await doc.loadInfo();

    let sheet: GoogleSpreadsheetWorksheet;
    if (typeof sheetNameOrIndex === "number") {
      if (
        sheetNameOrIndex >= 0 &&
        sheetNameOrIndex < doc.sheetsByIndex.length
      ) {
        console.log(`Índice de aba "${sheetNameOrIndex}" inválido.`);

        sheet = doc.sheetsByIndex[sheetNameOrIndex];
      } else {
        console.error(`Índice de aba "${sheetNameOrIndex}" inválido.`);
        return null;
      }
    } else if (typeof sheetNameOrIndex === "string") {
      const foundSheet = doc.sheetsByTitle[sheetNameOrIndex];
      if (foundSheet) {
        sheet = foundSheet;
      } else {
        console.error(`Aba com o nome "${sheetNameOrIndex}" não encontrada.`);
        return null;
      }
    } else {
      console.log(`Fallback para a primeira aba`);
      sheet = doc.sheetsByIndex[0]; // Fallback para a primeira aba
    }
    await sheet._ensureHeaderRowLoaded();

    const headers = sheet.headerValues;

    // Encontra o índice da coluna com o header especificado
    const columnIndex = headers.indexOf(columnHeader);

    if (columnIndex === -1) {
      console.error(
        `Coluna com o header "${columnHeader}" não encontrada na aba "${sheet.title}".`
      );
      return null;
    }

    const columnValues: any[] = [];
    const rows = await sheet.getRows();

    console.log(columnValues);

    for (const row of rows) {
      const isRowEmpty = Object.values(row).every(
        (value) => value === undefined || value === ""
      );
      if (isRowEmpty) {
        console.log("Linha vazia encontrada. Interrompendo a leitura.");
        break;
      }
      columnValues.push(row.get(columnHeader));
    }

    const missingInExpected = columnValues.filter(
      (value) => !expectedValues.includes(value)
    );
    const missingInSheet = expectedValues.filter(
      (value) => !columnValues.includes(value)
    );

    return { missingInExpected, missingInSheet };
  } catch (error: any) {
    console.error("Ocorreu um erro:", error.message);
    return null;
  }
}
