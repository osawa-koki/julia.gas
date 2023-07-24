/* eslint-disable @typescript-eslint/no-non-null-assertion */

const scriptProperties = PropertiesService.getScriptProperties()
const SPREADSHEET_FILE_ID = scriptProperties.getProperty('SPREADSHEET_FILE_ID')!
const SPREADSHEET_SHEET_NAME = scriptProperties.getProperty('SPREADSHEET_SHEET_NAME')!
const WIDTH = Number(scriptProperties.getProperty('WIDTH')!)
const HEIGHT = Number(scriptProperties.getProperty('HEIGHT')!)
const X_MIN = Number(scriptProperties.getProperty('X_MIN')!)
const X_MAX = Number(scriptProperties.getProperty('X_MAX')!)
const Y_MIN = Number(scriptProperties.getProperty('Y_MIN')!)
const Y_MAX = Number(scriptProperties.getProperty('Y_MAX')!)
const C_REAL = Number(scriptProperties.getProperty('C_REAL')!)
const C_IMAG = Number(scriptProperties.getProperty('C_IMAG')!)
const MAX_ITERATIONS = Number(scriptProperties.getProperty('MAX_ITERATIONS')!)
const CELL_SIZE = Number(scriptProperties.getProperty('CELL_SIZE')!)

const file = DriveApp.getFileById(SPREADSHEET_FILE_ID)
const sheet = SpreadsheetApp.open(file)
  .getSheetByName(SPREADSHEET_SHEET_NAME)!

function rgb2Hex (r: number, g: number, b: number): string {
  return '#' + ('0' + r.toString(16)).slice(-2) + ('0' + g.toString(16)).slice(-2) + ('0' + b.toString(16)).slice(-2)
}

function julia (): void {
  // 行数の取得。
  const lastRow = sheet.getLastRow()
  // 不足している行数を計算。
  const diffRow = HEIGHT - lastRow
  // 不足している行数がある場合は行を追加。
  if (diffRow > 0) sheet.insertRows(lastRow + 1, diffRow)

  // 列数の取得。
  const lastColumn = sheet.getLastColumn()
  // 不足している列数を計算。
  const diffColumn = WIDTH - lastColumn
  // 不足している列数がある場合は列を追加。
  if (diffColumn > 0) sheet.insertColumns(lastColumn + 1, diffColumn)

  // サイズを設定。
  for (let i = 0; i < WIDTH; i++) sheet.setColumnWidth(i + 1, CELL_SIZE)
  for (let i = 0; i < HEIGHT; i++) sheet.setRowHeight(i + 1, CELL_SIZE)

  // ジュリア集合の計算。
  for (let y = 0; y < HEIGHT; y++) {
    for (let x = 0; x < WIDTH; x++) {
      let zReal = X_MIN + (X_MAX - X_MIN) * x / WIDTH
      let zImag = Y_MIN + (Y_MAX - Y_MIN) * y / HEIGHT
      let i = 0
      for (; i < MAX_ITERATIONS; i++) {
        const zReal2 = zReal * zReal
        const zImag2 = zImag * zImag
        if (zReal2 + zImag2 > 4) break
        const zRealTemp = zReal2 - zImag2 + C_REAL
        const zImagTemp = 2 * zReal * zImag + C_IMAG
        zReal = zRealTemp
        zImag = zImagTemp
      }
      sheet.getRange(y + 1, x + 1).setBackground(rgb2Hex(0, 0, i * 255 / MAX_ITERATIONS))
    }
  }
}

export { julia }
