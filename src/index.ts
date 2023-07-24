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
const MAX_ITERATIONS = Number(scriptProperties.getProperty('MAX_ITERATIONS')!)
const CELL_SIZE = Number(scriptProperties.getProperty('CELL_SIZE')!)

const file = DriveApp.getFileById(SPREADSHEET_FILE_ID)
const sheet = SpreadsheetApp.open(file)
  .getSheetByName(SPREADSHEET_SHEET_NAME)!

function rgb2Hex (r: number, g: number, b: number): string {
  return '#' + ('0' + r.toString(16)).slice(-2) + ('0' + g.toString(16)).slice(-2) + ('0' + b.toString(16)).slice(-2)
}

function mandelbrot (): void {
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

  // マンデルブロ集合の計算。
  for (let y = 0; y < HEIGHT; y++) {
    for (let x = 0; x < WIDTH; x++) {
      let zx = 0
      let zy = 0
      let zx2 = 0
      let zy2 = 0
      let i = 0
      while (zx2 + zy2 <= 4 && i < MAX_ITERATIONS) {
        zy = 2 * zx * zy + (y / HEIGHT) * (Y_MAX - Y_MIN) + Y_MIN
        zx = zx2 - zy2 + (x / WIDTH) * (X_MAX - X_MIN) + X_MIN
        zx2 = zx * zx
        zy2 = zy * zy
        i++
      }
      sheet.getRange(y + 1, x + 1).setBackground(rgb2Hex(0, 0, i * 255 / MAX_ITERATIONS))
    }
  }
}

export { mandelbrot }
