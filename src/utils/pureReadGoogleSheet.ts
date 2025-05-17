import { google } from 'googleapis';
import { auth } from '../auth/googleAuth'; // Isso deve retornar um GoogleAuth configurado

export class PureSpreadsheetComparator {
  private sheetsApi;

  constructor(private spreadsheetId: string) {
    this.sheetsApi = google.sheets({ version: 'v4', auth });
  }

  public async compareColumnWithArray(
    columnHeader: string,
    expectedValues: any[],
    sheetNameOrIndex: string | number = 0
  ): Promise<null | Record<string, string[]>> {
    try {
      const sheetName = await this.resolveSheetName(sheetNameOrIndex);
      if (!sheetName) return null;

      const columnValues = await this.extractColumnValues(sheetName, columnHeader);
      if (!columnValues) return null;

      return this.compareValues(columnValues, expectedValues);
    } catch (error: any) {
      console.error('Ocorreu um erro:', error.message);
      return null;
    }
  }

  private async resolveSheetName(sheetNameOrIndex: string | number): Promise<string | null> {
    const response = await this.sheetsApi.spreadsheets.get({
      spreadsheetId: this.spreadsheetId,
    });

    const sheets = response.data.sheets;
    if (!sheets) return null;

    if (typeof sheetNameOrIndex === 'string') {
      const found = sheets.find((sheet) => sheet.properties?.title === sheetNameOrIndex);
      return found?.properties?.title || null;
    }

    if (typeof sheetNameOrIndex === 'number') {
      return sheets[sheetNameOrIndex]?.properties?.title || null;
    }

    return null;
  }

  private async extractColumnValues(sheetName: string, columnHeader: string): Promise<any[] | null> {
    const range = `${sheetName}!A:Z`; // Ajuste o range conforme necessário
    const response = await this.sheetsApi.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return null;

    const headers = rows[0];
    const columnIndex = headers.indexOf(columnHeader);
    if (columnIndex === -1) {
      console.error(`Coluna com o header "${columnHeader}" não encontrada.`);
      return null;
    }

    const columnValues: any[] = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const isRowEmpty = row.every((cell) => !cell || cell.trim() === '');
      if (isRowEmpty) break;

      columnValues.push(row[columnIndex]);
    }

    return columnValues;
  }

  private compareValues(sheetValues: any[], expectedValues: any[]) {
    const missingInExpected = sheetValues.filter((v) => !expectedValues.includes(v));
    const missingInSheet = expectedValues.filter((v) => !sheetValues.includes(v));

    return { missingInExpected, missingInSheet };
  }
}
