'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});

var _stream = require('stream');

var _templates = require('./templates');

/** Class representing a XLSX Row transformation from array to Row. */
class XLSXRowTransform extends _stream.Transform {
    constructor() {
        super({ objectMode: true });
        this.rowCount = 0;
    }
    /**
     * Transform array to row string
     */
    _transform(row, encoding, callback) {
        // eslint-disable-line
        if (!row || row.length == 0) return callback();

        const xlsxRow = (0, _templates.Row)(this.rowCount, row);
        this.rowCount++;
        callback(null, xlsxRow);
    }
}
exports.default = XLSXRowTransform;