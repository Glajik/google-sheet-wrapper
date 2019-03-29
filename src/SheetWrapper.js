import { SheetHelper } from '@airbag/sheet-helper';

/**
 * Provide more convient work with specified sheet
 * TODO:
 * 1. Write tests
 * 2. method updateSheet() maybe isn't work as expected
 */
// eslint-disable-next-line import/prefer-default-export
export class SheetWrapper extends SheetHelper {
  // eslint-disable-next-line no-useless-constructor
  constructor(options) {
    super(options);
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
    const values = this.sheet.getDataRange().getValues();
    return values;
  }

  /**
   * Get all sheet values, convert to collection of rowData objects
   * @returns array of rowData objects
   */
  get dataColl() {
    return this.toRowDataColl(this.values);
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
   * Insert row between header rows and first data row.
   * (universal method - both for rowData object or array)
   */
  insertRow(data) {
    this.sheet.insertRowBefore(this.firstRow);
    this.updateRow(this.firstRow, data);
    return this.firstRow;
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
    const range = this.getRowRange(rowId);
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
    const range = this.getRowRange(rowId);
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
    const row = this.numHeaders + 1;
    const column = 1;
    const numRows = sheet.getLastRow() - row + 1;
    if (numRows < 1) {
      return;
    }
    const numColumns = this.fields.length;
    sheet.getRange(row, column, numRows, numColumns).clearContent();
  }

  get headerValues() {
    return this.values.slice(0, this.numHeaders);
  }

  /**
   * Update all sheet except header rows
   * @param {Array} rowDataColl array of rowData objects
   */
  updateSheet(rowDataColl) {
    const values = super.toRowValuesColl(rowDataColl, this.headerValues);

    // update sheet
    const row = 1;
    const column = 1;
    const numColumns = this.fields.length;
    const numRows = values.length;
    this.sheet.getDataRange().clearContent();
    this.sheet.getRange(row, column, numRows, numColumns).setValues(values);
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
    const length = this.sheet.getLastRow() - this.firstRow;
    this.sheet.showRows(this.firstRow, length);
  }

  /**
   * find column id by name
   * @return index of column started from 1
   * @param {*} field field name
   */
  findColumnId(field) {
    return super.findColumnId(field);
  }
}
