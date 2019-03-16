import { expect } from 'chai';
// import * as sinon from 'sinon';

import { Sheet, Spreadsheet } from './SpreadsheetApp'; // eslint-disable-line
import SpreadsheetApp from './SpreadsheetApp'; // eslint-disable-line

// import after fake SpreadsheetApp
import SheetWrapper from '../src/sheetwrapper'; // eslint-disable-line

describe('Test', () => {
  before(() => {
  });

  it('Should init default', () => {
    const wrapped = new SheetWrapper({}, undefined, SpreadsheetApp);
    expect(wrapped).is.exist; // eslint-disable-line
  });

  it('Should custom init', () => {
    // eslint-disable-next-line
    const spreadsheetId = '12345';

    const cc = {
      sheetName: 'TestName',
      numHeaders: 1,
      fields: [
        'A',
        'B',
        'C',
      ],
    };

    // const fakeSheet = new Sheet();
    // sinon.stub(fakeSheet, 'getName').returns('TestName1');

    // const SpreadsheetApp = {
    //   getActiveSheet: () => fakeSheet,
    //   getActiveSpreadsheet: () => new Spreadsheet(),
    //   openById: id => new Spreadsheet(id),
    // };

    const wrapped = new SheetWrapper(cc, spreadsheetId, SpreadsheetApp);

    expect(wrapped.sheetName).is.equal(cc.sheetName);

    const sameSheet = SpreadsheetApp.openById(cc.spreadsheetId).getSheetByName(cc.sheetName);
    expect(wrapped.sheet, 'wrapped sheet').deep.equal(sameSheet, 'sameSheet');
  });
});
