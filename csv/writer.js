var inherits = require('util').inherits;

var prepareString = function (value) {
    return value
        .replace(/[\n\r]/g, ' ')
        .replace(/"/g, '""');
};

var Types = require('../types');
var BaseSpreadsheetWriter = require('../base');

//var toOADate = require('../utils').toOADate;
var toXMLDateTime = require('../utils').toXMLDateTime;
var stripHtml = require('../utils').stripHtml;

var CellType = Types.CellType;

/**
 * @class CsvSpreadsheetWriter
 * @extends BaseSpreadsheetWriter
 * */
var CsvSpreadsheetWriter = function () {
    this._init.apply(this, arguments);
};

inherits(CsvSpreadsheetWriter, BaseSpreadsheetWriter);

/**
 * @constructs
 * @param {Stream.Writable} stream - The output stream for the spreadsheet
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter}
 */
CsvSpreadsheetWriter.prototype._init = function (stream) {
    this._out = stream;

    /**
     * @private
     * @type string
     */
    this._encoding = 'utf8';

    /**
     * @private
     * @type boolean
     */
    this._shouldEndWorksheet = false;

    /**
     * @private
     * @type boolean
     */
    this._shouldEndRow = false;

    return this;
};

/**
 * Returns the appropriate mime type for this writer's output
 * @public
 * @this {CsvSpreadsheetWriter}
 * @returns {String}
 */
CsvSpreadsheetWriter.prototype.mimeType = function () {
    return 'application/vnd.ms-excel';
};

/**
 * Returns the appropriate file extension for this writer's output
 * @public
 * @this {CsvSpreadsheetWriter}
 * @returns {String}
 */
CsvSpreadsheetWriter.prototype.fileExtension = function () {
    return 'csv';
};

/**
 * Sets the output encoding. Default is utf8.
 * @param {String} encoding
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter}
 */
CsvSpreadsheetWriter.prototype.setEncoding = function (encoding) {
    this._encoding = encoding;
    return this;
};

/**
 * Gets the current output encoding. Default is utf8.
 * @this {CsvSpreadsheetWriter}
 * @returns {String}
 */
CsvSpreadsheetWriter.prototype.getEncoding = function () {
    return this._encoding;
};

/**
 * CSV does not support styling
 * @public
 * @param {?} style
 * @this {CsvSpreadsheetWriter}
 * @returns {int} - -1, as CSV does not support styling
 */
CsvSpreadsheetWriter.prototype.addStyle = function (style) {
    return -1;
};

/**
 * Adds a column.
 * <code>addColumn()</code> must be called for each column.
 * If cells are added which exceed the count of defined columns - then Excel programs will raise a schema error.
 * @public
 * @param {Number?} width - The width of this column. A width of zero means "automatic".
 * @param {boolean?} autoFitWidth - Should this column auto fit ts width
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.addColumn = function (width, autoFitWidth) {
    return this;
};

/**
 * Writes the beginning of the file
 * @public
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.beginFile = function () {
    return this;
};

/**
 * Writes the end of the file
 * @public
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.endFile = function () {

    if (this._endedFile) {
        return this;
    }

    this._endWorksheet();

    /**
     * @private
     * @type {boolean}
     */
    this._endedFile = true;

    setImmediate((function () {
        this.emit('finish');
    }).bind(this));

    return this;
};

/**
 * Begins a new worksheet
 * @public
 * @param {String?} worksheetName - The title for the new worksheet
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.newWorksheet = function (worksheetName) {

    var shouldAddEmptyRow = this._shouldEndWorksheet;

    this._endWorksheet();
    this._shouldEndWorksheet = true;

    if (shouldAddEmptyRow) {
        this.beginRow();
    }

    if (worksheetName != null) {
        this
            .beginRow()
            .addCell(worksheetName, CellType.String)
            .beginRow();
    }

    return this;
};

/**
 * Ends the current worksheet
 * @private
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype._endWorksheet = function () {

    if (this._shouldEndWorksheet) {
        this._endRow();
        this._shouldEndWorksheet = false;
    }

    return this;
};

/**
 * Begins a new row
 * @public
 * @param {int?} styleIndex - The index of the style to use.
 * @param {int?} height - Specific height for this row
 * @param {boolean?} autofitHeight - Should this row autofit its height
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.beginRow = function (styleIndex, height, autofitHeight) {

    if (!this._shouldEndWorksheet) {
        this.newWorksheet('New Sheet');
    }

    this._endRow();

    this._shouldEndRow = true;

    return this;
};

/**
 * Adds a new cell in the current row.
 * @public
 * @param {String|Number|Date|Boolean} data - Data to print out
 * @param {CellType?} type - The cell type
 * @param {int?} styleIndex - The style index to use for this cell
 * @param {int?} mergeAcross - How many cells should this cell span on?
 * @param {int?} mergeDown - How many rows should this cell span on?
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.addCell = function (data, type, styleIndex, mergeAcross, mergeDown) {

    var cellDef = '"';

    if (data instanceof Date) {
        cellDef += toXMLDateTime(data);
    } else if (data === true) {
        cellDef += '1';
    } else if (data === false) {
        cellDef += '0';
    } else {
        cellDef += prepareString(data + '');
    }

    cellDef += '",';
    this._out.write(cellDef, 'utf8');

    return this;
};

/**
 * Adds a new formula cell in the current row.
 * @public
 * @param {String} formula - The formula
 * @param {String|Number|Date|Boolean} dataPlaceholder - Placeholder data
 * @param {CellType} type - The cell type
 * @param {int?} styleIndex - The style index to use for this cell
 * @param {int?} mergeAcross - How many cells should this cell span on?
 * @param {int?} mergeDown - How many rows should this cell span on?
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.addFormulaCell = function (formula, dataPlaceholder, type, styleIndex, mergeAcross, mergeDown) {

    var cellDef = '"';

    if (dataPlaceholder instanceof Date) {
        cellDef += toXMLDateTime(dataPlaceholder);
    } else if (dataPlaceholder === true) {
        cellDef += '1';
    } else if (dataPlaceholder === false) {
        cellDef += '0';
    } else {
        cellDef += prepareString(dataPlaceholder + '');
    }

    cellDef += '",';
    this._out.write(cellDef, 'utf8');

    return this;
};

/**
 * Adds a new rich-text cell in the current row.
 * @public
 * @param {String} xml - Well-formed XML.
 * @param {int?} styleIndex - The style index to use for this cell
 * @param {int?} mergeAcross - How many cells should this cell span on?
 * @param {int?} mergeDown - How many rows should this cell span on?
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype.addRichTextCell = function (xml, styleIndex, mergeAcross, mergeDown) {

    var cellDef = '"';
    cellDef += stripHtml(xml + '');
    cellDef += '",';
    this._out.write(cellDef, 'utf8');

    return this;
};

/**
 * Writes the end of the current row
 * @private
 * @this {CsvSpreadsheetWriter}
 * @returns {CsvSpreadsheetWriter} <code>this</code>
 */
CsvSpreadsheetWriter.prototype._endRow = function () {

    if (this._shouldEndRow) {
        this._out.write('\n');
        this._shouldEndRow = false;
    }

    return this;
};

/**
 * @module
 * @type {CsvSpreadsheetWriter}
 * */
module.exports = CsvSpreadsheetWriter;