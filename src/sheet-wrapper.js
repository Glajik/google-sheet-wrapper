/* eslint-disable import/prefer-default-export */
/* eslint-disable no-underscore-dangle */

import findIndex from 'lodash/findIndex';

/**
 * Provide more convient work with specified sheet
 */
export class SheetWrapper {
  constructor(spreadsheetApi, spreadsheetId, sheetConfig = {}) {
    const {
      sheetName,
      numHeaders,
      fields,
    } = sheetConfig;

    // SpreadsheetApp is Google Api global object
    this.SpreadsheetApp = spreadsheetApi || SpreadsheetApp; // eslint-disable-line no-undef
    this.fields = fields;
    this.numHeaders = numHeaders || 0;
    this.spreadsheetId = spreadsheetId;
    this._sheetName = sheetName;
  }

  /**
   * @returns Spreadsheet object
   */
  get spreadsheet() {
    if (this._spreadsheet === undefined) {
      if (!this.spreadsheetId) {
        this._spreadsheet = this.SpreadsheetApp.getActiveSpreadsheet();
      } else {
        this._spreadsheet = this.SpreadsheetApp.openById(this.spreadsheetId);
      }
    }
    return this._spreadsheet;
  }

  /**
   * @returns Sheet object
   */
  get sheet() {
    if (this._sheet === undefined) {
      if (!this.sheetName) {
        this._sheet = this.spreadsheet.getActiveSheet();
      } else {
        this._sheet = this.spreadsheet.getSheetByName(this.sheetName);
      }
    }
    return this._sheet;
  }

  /**
   * @returns Name of sheet
   */
  get sheetName() {
    if (this._sheetName === undefined) {
      this._sheetName = this.sheet.getName();
    }
    return this._sheetName;
  }

  /**
   * @returns first data row number after headers
   */
  get firstRow() {
    return this.numHeaders + 1;
  }

  /**
   * Convert array of array of values to array of row object.
   * Each row contains the row index rowId started from 1.
   * @param {array} valuesColl array of row's arrays [[a, 1], [b, 2], ...]
   * @param {*} fields array of fields
   * @param {*} numHeaders count of header rows
   * @returns array of row objects [{ name: 'a', value: '1' }, ...]
   */
  static _convert(valuesColl, fields, numHeaders) {
    // remove header
    if (numHeaders) {
      for (let i = 0; i < numHeaders; i++) { // eslint-disable-line no-plusplus
        valuesColl.shift();
      }
    }

    const objectColl = [];
    const valuesCount = valuesColl.length;
    const fieldsCount = fields.length;

    for (let i = 0; i < valuesCount; i++) { // eslint-disable-line no-plusplus
      const obj = {};
      for (let j = 0; j < fieldsCount; j++) { // eslint-disable-line no-plusplus
        obj[fields[j]] = valuesColl[i][j];
      }
      obj.index = i;
      obj.rowId = i + 1 + numHeaders;
      objectColl.push(obj);
    }

    return objectColl;
  }

  /**
   * Load all data from sheet, and cache it
   * @returns array of rowData objects
   */
  get sheetData() {
    if (this._sheetData === undefined) {
      const dataRange = this.sheet.getDataRange();
      const valuesColl = dataRange.getValues();
      this._sheetData = SheetWrapper._convert(valuesColl, this.fields, this.numHeaders);
    }
    return this._sheetData;
  }

  /**
   * Next calling this.sheetData get data directly from sheet
   */
  resetSheetDataCache() {
    this._sheetData = undefined;
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
   * @returns rowData object
   * @param {Object} range range object
   */
  getDataFromRange(range) {
    const values = range.getValues()[0];
    const fn = (acc, field, index) => ({ ...acc, [field]: values[index] });
    return this.fields.reduce(fn, {});
  }

  /**
   * Inject rowId to returned object
   * @returns rowData object
   * @param {Number} rowId index of row, started from 1
   */
  getRowData(rowId) {
    const range = this.getRowRange(rowId);
    const data = this.getDataFromRange(range);
    return { ...data, rowId };
  }

  /**
   * Convert rowData object to array of row values
   * @return Array
   * @param {Object} rowData index of row, started from 1
   */
  getValues(rowData) {
    const fn = (acc, field) => {
      const value = rowData[field];
      return [...acc, value];
    };
    return this.fields.reduce(fn, []);
  }

  /**
   * find column id by name
   * @return index of column started from 1
   * @param {String} field FieldName
   */
  findColumnId(field) {
    return findIndex(this.fields, v => v === field) + 1;
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
      return this._appendRowArr(data);
    }
    if (data instanceof Object) {
      return this._appendRowObj(data);
    }
    return null;
  }

  /**
   * Append row to end of table - array version
   * @param {Array} values
   */
  _appendRowArr(values) {
    const rowId = this.sheet.appendRow(values);
    this.resetSheetDataCache();
    return rowId;
  }

  /**
   * Append row to end of table - rowData version
   * @param {Object} rowData
   */
  _appendRowObj(rowData) {
    const values = this.getValues(rowData);
    const rowId = this.sheet.appendRow(values);
    this.resetSheetDataCache();
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
      return this._updateRowObj(rowId, data);
    }
    if (data instanceof Array) {
      return this._updateRowArr(rowId, data);
    }
    this.resetSheetDataCache();
    return null;
  }

  /**
   * Update specified row with array of values.
   * @param {Number} rowId index of row, started from 1
   * @param {Array} values list of values
   */
  _updateRowArr(rowId, values) {
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
  _updateRowObj(rowId, rowData) {
    const numColumns = this.fields.length;
    const range = this.sheet.getRange(rowId, 1, 1, numColumns);
    const fields = Object.keys(rowData);

    // update by each field
    fields.forEach((field) => {
      const column = this.findColumnId(field);
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

    this._sheetData = undefined;
  }

  /**
   * Update all sheet except header rows
   * @param {Array} rowDataColl array of rowData objects
   */
  updateSheet(rowDataColl) {
    const { sheet } = this;

    this.clearSheet();

    const row = this.numHeaders + 1;
    const column = 1;
    const numColumns = this.fields.length;

    // convert rowData array to array of row values
    const values = rowDataColl.map(rowData => this.getValues(rowData));

    // update sheet
    const numRows = rowDataColl.length;
    sheet.getRange(row, column, numRows, numColumns).setValues(values);

    this.resetSheetDataCache();
    this.SpreadsheetApp.flush();
  }

  static _blockBuilder(acc, { rowId }) {
    const [first, ...rest] = acc;
    // init acc
    if (!first) {
      return [{ rowId, count: 1 }];
    }

    // if current rowId is sequence - modify count of first element
    const { count } = first;
    if (first.rowId + count === rowId) {
      return [{ rowId: first.rowId, count: count + 1 }, ...rest];
    }

    // sequence break - add new element
    return [{ rowId, count: 1 }, first, ...rest];
  }

  /**
   * Hide rows filtered by predicate function
   * @param {Function} predicate
   */
  hide(predicate) {
    const data = this.sheetData;
    const filtered = data.filter(predicate);
    this.SpreadsheetApp.getActiveSpreadsheet().toast('Start hiding');

    const blocks = filtered.reduce(SheetWrapper._blockBuilder, []);

    blocks.forEach(({ rowId, count }) => this.sheet.hideRows(rowId, count));
  }

  /**
   * Show rows filtered by predicate function
   * @param {Function} predicate
   */
  show(predicate) {
    const data = this.sheetData;
    const filtered = data.filter(predicate);
    this.SpreadsheetApp.getActiveSpreadsheet().toast('Start showing');

    const blocks = filtered.reduce(SheetWrapper._blockBuilder, []);
    blocks.forEach(({ rowId, count }) => this.sheet.showRows(rowId, count));
  }

  /** Show all hidden rows */
  showAll() {
    const length = this.sheet.getLastRow();
    this.sheet.showRows(3, length);
  }
}
