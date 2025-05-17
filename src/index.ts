import { SpreadsheetComparator } from "./utils/readGoogleSheet";

async function main() {
  const spreadsheetId = "1Z49bTwCdS6a_QG5gquooUIbUf4_8rbJNBBfd36hZI90"; // Substitua pela ID da sua planilha
  const columnHeader = "HEADAR_1"; // Substitua pelo header da coluna que você quer comparar
  const expectedValues = ["Alice", "Andrew", "Peter"]; // Array com os valores esperados (ordem diferente do exemplo anterior)
  const sheetNameOrIndex: string | number = "Sheet1"; // Ou o índice da aba (ex: 0 para a primeira aba)

  const comparator = new SpreadsheetComparator(spreadsheetId);
  const result = await comparator.compareColumnWithArray(
    columnHeader,
    expectedValues,
    sheetNameOrIndex
  );

  if (result) {
    if (
      result.missingInExpected.length === 0 &&
      result.missingInSheet.length === 0
    ) {
      console.log(
        `The column "${columnHeader}" in sheet "${sheetNameOrIndex}" has exactly the same values as the array.`
      );
    } else {
      console.log(
        `Differences found in column "${columnHeader}" in sheet "${sheetNameOrIndex}":`
      );
      if (result.missingInExpected.length > 0) {
        console.log(
          `  Values present in the spreadsheet, but not in the array: ${result.missingInExpected.join(
            ", "
          )}`
        );
      }
      if (result.missingInSheet.length > 0) {
        console.log(
          `  Values present in the array, but not in the spreadsheet: ${result.missingInSheet.join(
            ", "
          )}`
        );
      }
    }
  } else {
    console.error("An error occurred during the comparison.");
  }
}

main().catch(console.error);
