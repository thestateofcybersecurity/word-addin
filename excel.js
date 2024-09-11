/* global Excel, Word */

async function insertExcelTable() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1:D5");  // Modify this range based on the Excel table structure
    range.load("values");
    await context.sync();

    await Word.run(async (wordContext) => {
      const wordBody = wordContext.document.body;
      wordBody.insertTable(range.values.length, range.values[0].length, Word.InsertLocation.end, range.values);
      await wordContext.sync();
    });
  });
}