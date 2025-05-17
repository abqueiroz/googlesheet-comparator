import {
  GoogleSpreadsheet,
  GoogleSpreadsheetWorksheet,
} from "google-spreadsheet";
import { auth } from "../auth/googleAuth";

export class SpreadsheetComparator {
  private doc: GoogleSpreadsheet;
  private sheet: GoogleSpreadsheetWorksheet | null = null;

  constructor(private spreadsheetId: string) {
    this.doc = new GoogleSpreadsheet(spreadsheetId, auth);
  }

  public async compareColumnWithArray(
    columnHeader: string,
    expectedValues: any[],
    sheetNameOrIndex: string | number = 0
  ): Promise<null | Record<string, string[]>> {
    try {
      await this.loadSpreadsheetInfo();
      const sheetLoaded = await this.loadSheet(sheetNameOrIndex);
      if (!sheetLoaded) return null;

      const columnValues = await this.extractColumnValues(columnHeader);
      if (!columnValues) return null;

      const comparisonResult = this.compareValues(columnValues, expectedValues);
      return comparisonResult;
    } catch (error: any) {
      console.error("Ocorreu um erro:", error.message);
      return null;
    }
  }

  private async loadSpreadsheetInfo() {
    await this.doc.loadInfo();
  }

  private async loadSheet(sheetNameOrIndex: string | number): Promise<boolean> {
    if (typeof sheetNameOrIndex === "number") {
      if (
        sheetNameOrIndex >= 0 &&
        sheetNameOrIndex < this.doc.sheetsByIndex.length
      ) {
        this.sheet = this.doc.sheetsByIndex[sheetNameOrIndex];
      } else {
        console.error(`Índice de aba "${sheetNameOrIndex}" inválido.`);
        return false;
      }
    } else if (typeof sheetNameOrIndex === "string") {
      const foundSheet = this.doc.sheetsByTitle[sheetNameOrIndex];
      if (foundSheet) {
        this.sheet = foundSheet;
      } else {
        console.error(`Aba com o nome "${sheetNameOrIndex}" não encontrada.`);
        return false;
      }
    } else {
      console.log("Fallback para a primeira aba.");
      this.sheet = this.doc.sheetsByIndex[0];
    }

    await this.sheet._ensureHeaderRowLoaded();
    return true;
  }

  private async extractColumnValues(columnHeader: string): Promise<any[] | null> {
    if (!this.sheet) return null;

    const headers = this.sheet.headerValues;
    const columnIndex = headers.indexOf(columnHeader);

    if (columnIndex === -1) {
      console.error(
        `Coluna com o header "${columnHeader}" não encontrada na aba "${this.sheet.title}".`
      );
      return null;
    }

    const columnValues: any[] = [];
    const rows = await this.sheet.getRows();

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

    return columnValues;
  }

  private compareValues(sheetValues: any[], expectedValues: any[]) {
    const missingInExpected = sheetValues.filter(
      (value) => !expectedValues.includes(value)
    );
    const missingInSheet = expectedValues.filter(
      (value) => !sheetValues.includes(value)
    );

    return { missingInExpected, missingInSheet };
  }
}