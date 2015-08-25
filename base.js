
var util = require('util');

var unimplementedFunction = function () {
    throw new Error('This function is not implemented');
};

/**
 * @public
 * @class BaseSpreadsheetWriter
 * @extends {events.EventEmitter}
 * */
var BaseSpreadsheetWriter = function () { };

util.inherits(BaseSpreadsheetWriter, require('events').EventEmitter);

/**
 * Sets the output encoding. Default is utf8.
 * @param {String} encoding
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter}
 */
BaseSpreadsheetWriter.prototype.setEncoding = function (encoding) {
    return this;
};

/**
 * Gets the current output encoding. Default is utf8.
 * @this {BaseSpreadsheetWriter}
 * @returns {String}
 */
BaseSpreadsheetWriter.prototype.getEncoding = function () {
    return 'utf8';
};

/**
 * Returns the appropriate mime type for this writer's output
 * @public
 * @this {BaseSpreadsheetWriter}
 * @returns {String}
 */
BaseSpreadsheetWriter.prototype.mimeType = unimplementedFunction;

/**
 * Returns the appropriate file extension for this writer's output
 * @public
 * @this {BaseSpreadsheetWriter}
 * @returns {String}
 */
BaseSpreadsheetWriter.prototype.fileExtension = unimplementedFunction;

/**
 * Add a predefined style, to use later for styling cells/rows
 * @public
 * @param {SpreadsheetCellStyle} style
 * @this {BaseSpreadsheetWriter}
 * @returns {int} - The ID of the added style, to pass for styling cells/rows
 */
BaseSpreadsheetWriter.prototype.addStyle = unimplementedFunction;

/**
 * Adds a column.
 * <code>addColumn()</code> must be called for each column.
 * If cells are added which exceed the count of defined columns - then Excel programs will raise a schema error.
 * @public
 * @param {Number?} width - The width of this column. A width of zero means "automatic".
 * @param {boolean?} autoFitWidth - Should this column auto fit ts width
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.addColumn = unimplementedFunction;

/**
 * Writes the beginning of the file
 * @public
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.beginFile = unimplementedFunction;

/**
 * Writes the end of the file
 * @public
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.endFile = unimplementedFunction;

/**
 * Begins a new worksheet
 * @public
 * @param {String?} worksheetName - The title for the new worksheet
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.newWorksheet = unimplementedFunction;

/**
 * Begins a new row
 * @public
 * @param {int?} styleIndex - The index of the style to use.
 * @param {int?} height - Specific height for this row
 * @param {boolean?} autofitHeight - Should this row autofit its height
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.beginRow = unimplementedFunction;

/**
 * Adds a new cell in the current row.
 * @public
 * @param {String|Number|Date|Boolean} data - Data to print out
 * @param {CellType?} type - The cell type
 * @param {int?} styleIndex - The style index to use for this cell
 * @param {int?} mergeAcross - How many cells should this cell span on?
 * @param {int?} mergeDown - How many rows should this cell span on?
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.addCell = unimplementedFunction;


/**
 * Adds a new formula cell in the current row.
 * @public
 * @param {String} formula - The formula
 * @param {String|Number|Date|Boolean} dataPlaceholder - Placeholder data
 * @param {CellType} type - The cell type
 * @param {int?} styleIndex - The style index to use for this cell
 * @param {int?} mergeAcross - How many cells should this cell span on?
 * @param {int?} mergeDown - How many rows should this cell span on?
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.addFormulaCell = unimplementedFunction;

/**
 * Adds a new rich-text cell in the current row.
 * @public
 * @param {String} xml - Well-formed XML.
 * @param {int?} styleIndex - The style index to use for this cell
 * @param {int?} mergeAcross - How many cells should this cell span on?
 * @param {int?} mergeDown - How many rows should this cell span on?
 * @this {BaseSpreadsheetWriter}
 * @returns {BaseSpreadsheetWriter} <code>this</code>
 */
BaseSpreadsheetWriter.prototype.addRichTextCell = unimplementedFunction;

/**
 * @module
 * @type {BaseSpreadsheetWriter}
 * */
module.exports = BaseSpreadsheetWriter;