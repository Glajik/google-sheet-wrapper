import { expect } from 'chai';


// is important to create stub before import SheetWrapper

// eslint-disable-next-line no-unused-vars
import SpreadsheetApp from './SpreadsheetApp';

// eslint-disable-next-line import/first
import { SheetWrapper } from '../src/SheetWrapper';

describe('Test', () => {
  before(() => {
  });

  it('Should init', () => {
    const wrapped = new SheetWrapper();
    expect(wrapped).is.exist; // eslint-disable-line
  });

  it('test getter SheetWrapper.values', () => {
    const options = {
      sheetName: 'TestName',
      numHeaders: 1,
      fields: 'A, B, C',
    };

    const values = [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
    ];

    // const fakeSheet = new Sheet();
    // sinon.stub(fakeSheet, 'getName').returns('TestName1');

    // const SpreadsheetApp = {
    //   getActiveSheet: () => fakeSheet,
    //   getActiveSpreadsheet: () => new Spreadsheet(),
    //   openById: id => new Spreadsheet(id),
    // };

    const sw = new SheetWrapper(options);
    console.log(sw.values);
  });
});
