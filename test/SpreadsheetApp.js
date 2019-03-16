/**
 * Class for moking Sheet object
 */
export class Sheet {
  constructor(name) {
    this.name = name || 'Sheet 1';
  }
  // eslint-disable-next-line
  getName() {
    return this.name;
  }
}

/**
 * Class for mocking Spreadsheet object
 */
export class Spreadsheet {
  constructor(id) {
    this.id = id;
  }

  getId() {
    return this.id;
  }

  // eslint-disable-next-line
  getActiveSheet() {
    return new Sheet();
  }

  // eslint-disable-next-line
  getSheetByName(name) {
    return new Sheet(name);
  }
}

/**
 * Export static methods
 */
// eslint-disable-next-line
const SpreadsheetApp = {
  getActiveSheet: () => new Sheet(),
  getActiveSpreadsheet: () => new Spreadsheet(),
  openById: id => new Spreadsheet(id),
};

export default SpreadsheetApp;
