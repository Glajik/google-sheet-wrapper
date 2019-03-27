import { SheetHelper } from '@airbag/sheet-helper';

/**
 * Provide more convient work with specified sheet
 * TODO:
 * 1. Write tests
 * 2. method updateSheet() maybe isn't work as expected
 */
// eslint-disable-next-line import/prefer-default-export
export class SheetWrapper extends SheetHelper {
  constructor(options) {
    super(options);
    this.memo = {
      values: undefined,
      dataColl: undefined,
    };
  }

  // eslint-disable-next-line class-methods-use-this
  get spreadsheet() {
    // eslint-disable-next-line no-undef
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  get sheet() {
    return this.spreadsheet.getSheetByName(this.sheetName);
  }

  get values() {
    if (this.memo.values === undefined) {
      this.memo.values = this.sheet.getDataRange().getValues();
    }
    return this.memo.values;
  }

  /**
   * Get all sheet values, convert to collection of rowData objects
   * @returns array of rowData objects
   */
  get dataColl() {
    if (this.memo.dataColl === undefined) {
      this.memo.dataColl = this.toRowDataColl(this.values);
    }
    return this.memo.dataColl;
  }

  /**
   * Next calling this.sheetData get data directly from sheet
   */
  reset() {
    this.memo = {
      values: undefined,
      dataColl: undefined,
    };
  }

  /**
   * @returns range object
   * @param {Number} rowId index of row started from 1
   */
  getRowRange(rowId) {
    const { sheet, fields } = this;
    const numColumns = fields.length;
    return sheet.getRange(rowId, 1, 1, numColumns);
  }

  /**
   * Inject rowId to returned object
   * @returns rowData object
   * @param {Number} rowId index of row, started from 1
   */
  getRowData(rowId) {
    const range = this.getRowRange(rowId);
    const [rowValues] = range.getValues();
    const data = super.toRowData(rowValues);
    return { ...data, rowId };
  }

  /**
   * @returns index of row with selected cell
   */
  getSelectedRow() {
    return this.sheet.getActiveCell().getRowIndex();
  }

  /**
   * Insert row between header rows and first data row.
   * (universal method - both for rowData object or array)
   */
  insertRow(data) {
    this.sheet.insertRowBefore(this.firstRow);
    this.resetSheetDataCache();
    this.updateRow(this.firstRow, data);
    return this.firstRow;
  }

  /**
   * Apend row to end of table.
   * (universal method - both for rowData object or array)
   * @param {any} data rowData object or Array of row values
   */
  appendRow(data) {
    if (data instanceof Array) {
      return this.appendRowArr(data);
    }
    if (data instanceof Object) {
      return this.appendRowObj(data);
    }
    return null;
  }

  /**
   * Append row to end of table - array version
   * @param {Array} values
   */
  appendRowArr(values) {
    const rowId = this.sheet.appendRow(values);
    this.reset();
    return rowId;
  }

  /**
   * Append row to end of table - rowData version
   * @param {Object} rowData
   */
  appendRowObj(rowData) {
    const values = super.toRowValues(rowData);
    const rowId = this.sheet.appendRow(values);
    this.reset();
    return rowId;
  }

  /**
   * Update specified row.
   * (universal method - both for rowData object or array)
   * In case rowData object - you can use part of whole row data
   * and only this part be updated
   * @param {Number} rowId index of row, started from 1
   * @param {any} data rowData or array of row values
   */
  updateRow(rowId, data) {
    if (data instanceof Object) {
      return this.updateRowObj(rowId, data);
    }
    if (data instanceof Array) {
      return this.updateRowArr(rowId, data);
    }
    return null;
  }

  /**
   * Update specified row with array of values.
   * @param {Number} rowId index of row, started from 1
   * @param {Array} values list of values
   */
  updateRowArr(rowId, values) {
    const range = this.sheet.getRange(rowId, 1, 1, values.length);
    range.setFontWeight(null);
    range.setValues([values]);
    return range;
  }

  /**
   * Update specified row with rowData object.
   * rowData object can consist part of whole row data
   * and only this part be updated
   * @param {Number} rowId index of row, started from 1
   * @param {Object} rowData rowData object
   */
  updateRowObj(rowId, rowData) {
    const numColumns = super.fields.length;
    const range = this.sheet.getRange(rowId, 1, 1, numColumns);
    const fields = Object.keys(rowData);

    // update by each field
    fields.forEach((field) => {
      const column = super.findColumnId(field);
      if (column <= 0) {
        return;
      }
      range.getCell(1, column).setValue(rowData[field]);
    });

    return range;
  }

  /**
   * Clear sheet but leave headers
   */
  clearSheet() {
    const { sheet } = this;
    const row = super.numHeaders + 1;
    const column = 1;
    const numRows = sheet.getLastRow() - row + 1;
    if (numRows < 1) {
      return;
    }
    const numColumns = super.fields.length;
    sheet.getRange(row, column, numRows, numColumns).clearContent();

    this.reset();
  }

  /**
   * Update all sheet except header rows
   * @param {Array} rowDataColl array of rowData objects
   */
  updateSheet(rowDataColl) {
    const values = super.toRowValuesColl(rowDataColl);

    // update sheet
    const row = super.numHeaders + 1;
    const column = 1;
    const numColumns = super.fields.length;
    const numRows = rowDataColl.length;
    this.sheet
      .getRange(row, column, numRows, numColumns)
      .clearContent()
      .setValues(values);

    this.memo.dataColl = super.clone(rowDataColl);
    // eslint-disable-next-line no-undef
    SpreadsheetApp.flush();
  }

  /**
   * Hide rows filtered by predicate function
   * @param {Function} predicate
   */
  hide(predicate) {
    this.spreadsheet.toast('Start hiding');
    const blocks = super.getBlocks(this.dataColl, predicate);
    blocks.forEach(({ rowId, count }) => this.sheet.hideRows(rowId, count));
  }

  /**
   * Show rows filtered by predicate function
   * @param {Function} predicate
   */
  show(predicate) {
    this.spreadsheet.getActiveSpreadsheet().toast('Start showing');
    const blocks = super.getBlocks(this.dataColl, predicate);
    blocks.forEach(({ rowId, count }) => this.sheet.showRows(rowId, count));
  }

  /** Show all hidden rows */
  showAll() {
    const length = this.sheet.getLastRow() - super.firstRow;
    this.sheet.showRows(super.firstRow, length);
  }
}
