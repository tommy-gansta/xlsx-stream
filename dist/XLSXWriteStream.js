'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});

var _archiver = require('archiver');

var _archiver2 = _interopRequireDefault(_archiver);

var _stream = require('stream');

var _templates = require('./templates');

var templates = _interopRequireWildcard(_templates);

var _XLSXRowTransform = require('./XLSXRowTransform');

var _XLSXRowTransform2 = _interopRequireDefault(_XLSXRowTransform);

function _interopRequireWildcard(obj) { if (obj && obj.__esModule) { return obj; } else { var newObj = {}; if (obj != null) { for (var key in obj) { if (Object.prototype.hasOwnProperty.call(obj, key)) newObj[key] = obj[key]; } } newObj.default = obj; return newObj; } }

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

/** Class representing a XLSX Write Stream. */
class XLSXWriteStream {
    /**
     * Create new Stream
     */
    constructor() {
        this.zip = (0, _archiver2.default)('zip', {
            forceUTC: true
        });
        this.zip.catchEarlyExitAttached = true;

        this.zip.append(templates.ContentTypes, {
            name: '[Content_Types].xml'
        });

        this.zip.append(templates.Rels, {
            name: '_rels/.rels'
        });

        this.zip.append(templates.Workbook, {
            name: 'xl/workbook.xml'
        });

        this.zip.append(templates.Styles, {
            name: 'xl/styles.xml'
        });

        this.zip.append(templates.WorkbookRels, {
            name: 'xl/_rels/workbook.xml.rels'
        });

        this.zip.on('warning', err => {
            console.warn(err);
        });

        this.zip.on('error', err => {
            console.error(err);
        });

        this.finalize = this.finalize.bind(this);
    }

    setInputStream(stream) {
        const toXlsxRow = new _XLSXRowTransform2.default();
        const transformedStream = stream.pipe(toXlsxRow);
        this.sheetStream = new _stream.PassThrough();
        this.sheetStream.write(templates.SheetHeader);
        transformedStream.on('end', this.finalize);
        // stream.on('data', () => console.log('Input stream data'));
        transformedStream.pipe(this.sheetStream);
        this.zip.append(this.sheetStream, {
            name: 'xl/worksheets/sheet1.xml'
        });
    }

    getOutputStream() {
        return this.zip;
    }
    /**
     * Finalize the zip archive
     */
    finalize() {
        this.sheetStream.end(templates.SheetFooter);
        return this.zip.finalize();
    }
}
exports.default = XLSXWriteStream;