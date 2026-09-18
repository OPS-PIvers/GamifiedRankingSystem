/**
 * A minimal in-memory stand-in for the Apps Script services Code.js uses, so the submission and
 * point-calculation logic can be exercised with plain `node` instead of by hand in a live
 * spreadsheet.
 *
 * It is deliberately small: it models cells, formulas, appendRow and getLastRow closely enough to
 * catch off-by-one and ordering mistakes, and stubs everything else (mail, locks, the UI) just
 * enough to let the code run. It is not an emulator - anything it does not model is out of scope
 * for these tests.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CODE_JS = process.env.CODE_JS || path.join(__dirname, '..', 'Code.js');

/** An in-memory sheet: a sparse grid of values plus a separate map of formulas. */
class Sheet {
  constructor(name, rows = []) {
    this.name = name;
    this.rows = rows;
    this.formulas = {};
    this.protections = [];
  }

  getName() { return this.name; }

  /** Reads a cell by 1-indexed row/column, returning '' for anything unset. */
  _cell(row, col) {
    const r = this.rows[row - 1] || [];
    const v = r[col - 1];
    return v === undefined ? '' : v;
  }

  /** Writes a cell by 1-indexed row/column, growing the grid as needed. */
  _set(row, col, value) {
    while (this.rows.length < row) this.rows.push([]);
    const r = this.rows[row - 1];
    while (r.length < col) r.push('');
    r[col - 1] = value;
  }

  /** Mirrors Sheets: the last row holding any non-empty value. */
  getLastRow() {
    let last = 0;
    this.rows.forEach((row, i) => {
      if (row && row.some(v => v !== '' && v !== undefined && v !== null)) last = i + 1;
    });
    return last;
  }

  getLastColumn() {
    return Math.max(0, ...this.rows.map(r => (r ? r.length : 0)));
  }

  getRange(a, col, numRows, numCols) {
    if (typeof a === 'string') return this._getRangeByA1(a);

    const sheet = this;
    // Apps Script throws on a non-positive row/column/size; reproduce that so the tests catch
    // range arithmetic that goes out of bounds.
    if (a < 1 || col < 1 || numRows < 1 || numCols < 1) {
      throw new Error(`Range not found / invalid arguments: getRange(${a}, ${col}, ${numRows}, ${numCols}) on "${sheet.name}"`);
    }

    return {
      getRow: () => a,
      getColumn: () => col,
      getNumRows: () => numRows,
      getNumColumns: () => numCols,
      getSheet: () => sheet,
      getValue: () => sheet._cell(a, col),
      getValues() {
        const out = [];
        for (let i = 0; i < numRows; i++) {
          const row = [];
          for (let j = 0; j < numCols; j++) row.push(sheet._cell(a + i, col + j));
          out.push(row);
        }
        return out;
      },
      getFormulas() {
        const out = [];
        for (let i = 0; i < numRows; i++) {
          const row = [];
          for (let j = 0; j < numCols; j++) row.push(sheet.formulas[`${a + i},${col + j}`] || '');
          out.push(row);
        }
        return out;
      },
      setValue(v) { sheet._set(a, col, v); return this; },
      setValues(values) {
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) sheet._set(a + i, col + j, values[i][j]);
        }
        return this;
      },
      setFormula(formula) {
        sheet.formulas[`${a},${col}`] = formula;
        // A formula cell is not empty, so give it a value for getLastRow()'s benefit.
        sheet._set(a, col, 0);
        return this;
      },
      setFormulas(formulas) {
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) {
            sheet.formulas[`${a + i},${col + j}`] = formulas[i][j];
            if (sheet._cell(a + i, col + j) === '') sheet._set(a + i, col + j, 0);
          }
        }
        return this;
      },
      setFontWeight() { return this; },
      protect() {
        const protection = {
          description: '',
          getDescription() { return this.description; },
          setDescription(d) { this.description = d; return this; },
          setWarningOnly() { return this; },
          remove() { sheet.protections = sheet.protections.filter(p => p !== protection); },
        };
        sheet.protections.push(protection);
        return protection;
      },
    };
  }

  /** Supports the "D2" and "D2:E" forms used by the roster code. */
  _getRangeByA1(a1) {
    const m = a1.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d*))?$/);
    if (!m) throw new Error(`Unsupported A1 notation in mock: ${a1}`);
    const toCol = letters => letters.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);

    const firstCol = toCol(m[1]);
    const firstRow = parseInt(m[2], 10);
    const lastCol = m[3] ? toCol(m[3]) : firstCol;
    // An open-ended range like "D2:E" runs to the last populated row.
    const lastRow = m[4] ? parseInt(m[4], 10) : Math.max(firstRow, this.getLastRow());

    return this.getRange(firstRow, firstCol, lastRow - firstRow + 1, lastCol - firstCol + 1);
  }

  getDataRange() {
    return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
  }

  appendRow(row) {
    const target = this.getLastRow() + 1;
    row.forEach((value, i) => this._set(target, i + 1, value));
  }

  getProtections() { return this.protections.slice(); }

  clear() { this.rows = []; this.formulas = {}; }
}

/**
 * Loads Code.js into a fresh sandbox wired to the given sheets.
 *
 * @param {Object} sheets Map of sheet name to Sheet.
 * @param {Object} [options]
 * @param {string} [options.activeUser] What Session.getActiveUser().getEmail() returns.
 * @param {boolean} [options.lockFails] Make LockService.waitLock() throw, as it does under load.
 * @returns {Object} The sandbox. Code.js's functions are properties on it, and `sentEmails`
 *                   collects everything passed to MailApp.sendEmail.
 */
function loadCode(sheets, options = {}) {
  const sentEmails = [];

  const sandbox = {
    console: { log: () => {}, error: () => {} },
    Date, Math, String, Number, JSON, isNaN,
    sentEmails,

    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => sheets[name] || null,
        insertSheet: name => (sheets[name] = new Sheet(name)),
      }),
      flush: () => {},
      ProtectionType: { RANGE: 'RANGE' },
      getUi: () => ({
        alert: () => {},
        createMenu: () => ({ addItem() { return this; }, addToUi() {} }),
        ButtonSet: { YES_NO: 1 },
        Button: { YES: 1 },
      }),
    },

    LockService: {
      getScriptLock: () => ({
        waitLock: () => { if (options.lockFails) throw new Error('Could not obtain lock'); },
        releaseLock: () => {},
      }),
    },

    Session: { getActiveUser: () => ({ getEmail: () => options.activeUser || '' }) },

    MailApp: { sendEmail: message => sentEmails.push(message) },

    HtmlService: {
      createTemplateFromFile: () => ({ evaluate: () => ({ getContent: () => '<html></html>' }) }),
    },

    ScriptApp: {
      getProjectTriggers: () => [],
      newTrigger: () => ({ forSpreadsheet() { return this; }, onEdit() { return this; }, create() {} }),
      deleteTrigger: () => {},
    },
  };

  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(CODE_JS, 'utf8'), sandbox, { filename: CODE_JS });
  return sandbox;
}

module.exports = { Sheet, loadCode, CODE_JS };
