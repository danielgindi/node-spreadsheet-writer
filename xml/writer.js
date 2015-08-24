var Stream = require('stream');
var Color = require('color');
var inherits = require('util').inherits;

var prepareString = function (value) {
    return value
        .replace(/&/g, '&amp;"')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;')
        .replace(/\r/g, '&#xD;')
        .replace(/\n/g, '&#xA;');
};

var colorForSpreadsheet = function (color) {
    if (!color) return '';
    if (color.alpha() <= 0) return 'transparent';
    return color.hexString();
};

var Types = require('../types');
var BaseSpreadsheetWriter = require('../base');

//var toOADate = require('../utils').toOADate;
var toXMLDateTime = require('../utils').toXMLDateTime;

var VerticalAlignment = Types.VerticalAlignment;
var BorderLineStyle = Types.BorderLineStyle;
var BorderPosition = Types.BorderPosition;
var FontFamily = Types.FontFamily;
var FontUnderline = Types.FontUnderline;
var FontVerticalAlign = Types.FontVerticalAlign;
var HorizontalAlignment = Types.HorizontalAlignment;
var HorizontalReadingOrder = Types.HorizontalReadingOrder;
var InteriorPattern = Types.InteriorPattern;
//var NumberFormats = Types.NumberFormats;
var CellType = Types.CellType;
//var CellFormulaPlaceholder = Types.CellFormulaPlaceholder;

/**
 * @class XmlSpreadsheetWriter
 * @extends BaseSpreadsheetWriter
 * */
var XmlSpreadsheetWriter = function () {
    this._init.apply(this, arguments);
};

inherits(XmlSpreadsheetWriter, BaseSpreadsheetWriter);

/**
 * @constructs
 * @param {Stream.Writable} stream - The output stream for the spreadsheet
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter}
 */
XmlSpreadsheetWriter.prototype._init = function (stream) {
    this._out = stream;

    /**
     * @private
     * @type {Array.<SpreadsheetCellStyle>}
     *
     */
    this._styles = [];

    /**
     * @private
     * @type {Array.<{width: Number?, autoFitWidth: boolean?}>}
     *
     */
    this._columnWidths = [];

    /**
     * @private
     * @type boolean
     */
    this._shouldBeginWorksheet = false;

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

    /**
     * @private
     * @type boolean
     */
    this._hasWrittenStyles = false;

    return this;
};

/**
 * Returns the appropriate mime type for this writer's output
 * @public
 * @this {XmlSpreadsheetWriter}
 * @returns {String}
 */
XmlSpreadsheetWriter.prototype.mimeType = function () {
    return 'text/xml';
};

/**
 * Returns the appropriate file extension for this writer's output
 * @public
 * @this {XmlSpreadsheetWriter}
 * @returns {String}
 */
XmlSpreadsheetWriter.prototype.fileExtension = function () {
    return 'xml';
};

/**
 * Add a predefined style, to use later for styling cells/rows
 * @public
 * @param {SpreadsheetCellStyle} style
 * @this {XmlSpreadsheetWriter}
 * @returns {int} - The ID of the added style, to pass for styling cells/rows
 */
XmlSpreadsheetWriter.prototype.addStyle = function (style) {
    if (this._hasWrittenStyles) {
        throw new Error('Cannot add style at this phase. addStyle() must be called before writing data to the sheet');
    }
    this._styles.push(style);
    return this._styles.length - 1;
};

/**
 * Adds a column.
 * <code>addColumn()</code> must be called for each column.
 * If cells are added which exceed the count of defined columns - then Excel programs will raise a schema error.
 * @public
 * @param {Number?} width - The width of this column. A width of zero means "automatic".
 * @param {boolean?} autoFitWidth - Should this column auto fit ts width
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.addColumn = function (width, autoFitWidth) {
    if (!this._shouldBeginWorksheet) {
        throw new Error('Cannot add a column at this phase. Columns must be added after a call to newWorksheet(), and before adding any rows.');
    }
    if (typeof width === 'boolean') {
        autoFitWidth = /** @type boolean */width;
        width = null;
    }
    this._columnWidths.push({ width: width, autoFitWidth: autoFitWidth });
    return this;
};

/**
 * Writes the beginning of the file
 * @public
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.beginFile = function () {

    this._out.write(
        '<?xml version="1.0" encoding="utf-8"?>\n' +
        '<?mso-application progid="Excel.Sheet"?>\n' +
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
        ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
        ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
        ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n',
        'utf8');

    return this;
};

/**
 * Writes the end of the file
 * @public
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.endFile = function () {

    if (this._endedFile) {
        return this;
    }

    this._endWorksheet();
    this._out.write('</Workbook>\n');

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
 * Writes the styles to the beginning of the file
 * @private
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype._writeStyles = function () {

    if (!this._hasWrittenStyles) {

        this._out.write(
            ' <Styles>\n' +
            '  <Style ss:ID="Default" ss:Name="Normal">\n' +
            '   <Alignment ss:Vertical="Bottom"/>\n' +
            '  </Style>\n');

        for (var i = 0; i < this._styles.length; i++) {
            this.writeStyle(i);
        }

        this._out.write(' </Styles>\n');

        this._hasWrittenStyles = true;
    }

    return this;
};

/**
 * Writes the beginning of the new worksheet
 * @private
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype._beginWorksheet = function () {

    if (this._shouldBeginWorksheet) {
        this._out.write('  <Table ss:ExpandedColumnCount="' + this._columnWidths.length + '">\n');
        for (var i = 0; i < this._columnWidths.length; i++) {
            var column = this._columnWidths[i];

            this._out.write('   <Column');

            if (column.width != null) {
                this._out.write(' ss:Width="' + (column.width || 0) + '"');
            }

            if (typeof column.autoFitWidth === 'boolean') {
                this._out.write(' ss:AutoFitWidth="' + (column.autoFitWidth ? 1 : 0) + '"');
            }

            this._out.write('/>\n');
        }
        this._shouldBeginWorksheet = false;
    }

    return this;
};

/**
 * Begins a new worksheet
 * @public
 * @param {String?} worksheetName - The title for the new worksheet
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.newWorksheet = function (worksheetName) {

    if (this._shouldEndWorksheet) {
        this._endWorksheet();
    }
    else if (!this._hasWrittenStyles) {
        this._writeStyles();
    }

    this._shouldEndWorksheet = true;
    this._out.write(' <Worksheet ss:Name="' + (worksheetName == null ? '' : prepareString(worksheetName + '')) + '">\n', 'utf8');
    this._shouldBeginWorksheet = true;

    return this;
};

/**
 * Ends the current worksheet
 * @private
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype._endWorksheet = function () {

    if (this._shouldEndWorksheet) {
        this._endRow();
        if (!this._shouldBeginWorksheet) {
            this._out.write('  </Table>\n');
        }
        this._out.write(' </Worksheet>\n');
        this._shouldBeginWorksheet = false;
        this._shouldEndWorksheet = false;
    }

    this._columnWidths.length = 0;

    return this;
};

/**
 * Begins a new row
 * @public
 * @param {int?} styleIndex - The index of the style to use.
 * @param {int?} height - Specific height for this row
 * @param {boolean?} autofitHeight - Should this row autofit its height
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.beginRow = function (styleIndex, height, autofitHeight) {

    if (!this._shouldEndWorksheet) {
        this.newWorksheet('New Sheet');
    }

    this._beginWorksheet();
    this._endRow();

    this._out.write('   <Row');
    if (styleIndex != null && styleIndex >= 0) {
        this._out.write(' ss:StyleID="s' + (styleIndex + 21) + '"');
    }
    if (height != null) {
        this._out.write(' ss:Height="' + height + '"');
    }
    if (typeof autofitHeight === 'boolean') {
        this._out.write(' ss:AutoFitHeight="' + (autofitHeight ? 1 : 0) + '"');
    }
    this._out.write('>\n');

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
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.addCell = function (data, type, styleIndex, mergeAcross, mergeDown) {

    var cellDef = '    <Cell';

    if (styleIndex != null && styleIndex >= 0) {
        cellDef += ' ss:StyleID="s' + (styleIndex + 21) + '"';
    }

    if (mergeAcross > 0) {
        cellDef += ' ss:MergeAcross="' + mergeAcross + '"';
    }

    if (mergeDown > 0) {
        cellDef += ' ss:MergeDown="' + mergeDown + '"';
    }

    if (!type) {
        var dataType = typeof data;
        if (dataType === 'number') {
            type = CellType.Number;
        }
        else if (dataType === 'boolean') {
            type = CellType.Boolean;
        }
        else if (data instanceof Date) {
            type = CellType.DateTime;
        }
        else {
            type = CellType.String;
        }
    }

    cellDef += '><Data ss:Type="' + type + '">';

    if (data instanceof Date) {
        cellDef += toXMLDateTime(data);
    } else if (data === true) {
        cellDef += '1';
    } else if (data === false) {
        cellDef += '0';
    } else {
        cellDef += prepareString(data + '');
    }

    cellDef += '</Data></Cell>\n';
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
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.addFormulaCell = function (formula, dataPlaceholder, type, styleIndex, mergeAcross, mergeDown) {

    var cellDef = '    <Cell';

    if (styleIndex != null && styleIndex >= 0) {
        cellDef += ' ss:StyleID="s' + (styleIndex + 21) + '"';
    }

    if (mergeAcross > 0) {
        cellDef += ' ss:MergeAcross="' + mergeAcross + '"';
    }

    if (mergeDown > 0) {
        cellDef += ' ss:MergeDown="' + mergeDown + '"';
    }

    cellDef += ' ss:Formula="' + prepareString(formula + '') + '"';

    cellDef += '><Data ss:Type="' + (type || CellType.String) + '">';

    if (dataPlaceholder instanceof Date) {
        cellDef += toXMLDateTime(dataPlaceholder);
    } else if (dataPlaceholder === true) {
        cellDef += '1';
    } else if (dataPlaceholder === false) {
        cellDef += '0';
    } else {
        cellDef += prepareString(dataPlaceholder + '');
    }

    cellDef += '</Data></Cell>\n';
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
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.addRichTextCell = function (xml, styleIndex, mergeAcross, mergeDown) {

    var cellDef = '    <Cell';

    if (styleIndex != null && styleIndex >= 0) {
        cellDef += ' ss:StyleID="s' + (styleIndex + 21) + '"';
    }

    if (mergeAcross > 0) {
        cellDef += ' ss:MergeAcross="' + mergeAcross + '"';
    }

    if (mergeDown > 0) {
        cellDef += ' ss:MergeDown="' + mergeDown + '"';
    }

    cellDef += '><Data ss:Type="' + CellType.String + '">';
    cellDef += xml + '';
    cellDef += '</Data></Cell>\n';
    this._out.write(cellDef, 'utf8');

    return this;
};

/**
 * Writes the end of the current row
 * @private
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype._endRow = function () {

    if (this._shouldEndRow) {
        this._out.write('   </Row>\n');
        this._shouldEndRow = false;
    }

    return this;
};

/**
 * Renders the specified style to the output stream
 * @private
 * @param {int} styleIndex
 * @this {XmlSpreadsheetWriter}
 * @returns {XmlSpreadsheetWriter} <code>this</code>
 */
XmlSpreadsheetWriter.prototype.writeStyle = function (styleIndex) {
    var style = this._styles[styleIndex];

    this._out.write('  <Style ss:ID="s' + (styleIndex + 21) + '"');
    if (style.parentStyleIndex != null) {
        this._out.write(' ss:Parent="' + style.parentStyleIndex + '"', 'utf8');
    }
    this._out.write('>\n');

    if (style.alignment) {
        this._out.write('   <ss:Alignment'); // Opening tag

        var horizontal = null;
        switch (style.alignment.horizontal) {
            default:
            case HorizontalAlignment.Automatic: // Default
                break;
            case HorizontalAlignment.Left:
                horizontal = 'Left';
                break;
            case HorizontalAlignment.Center:
                horizontal = 'Center';
                break;
            case HorizontalAlignment.Right:
                horizontal = 'Right';
                break;
            case HorizontalAlignment.Fill:
                horizontal = 'Fill';
                break;
            case HorizontalAlignment.Justify:
                horizontal = 'Justify';
                break;
            case HorizontalAlignment.CenterAcrossSelection:
                horizontal = 'CenterAcrossSelection';
                break;
            case HorizontalAlignment.Distributed:
                horizontal = 'Distributed';
                break;
            case HorizontalAlignment.JustifyDistributed:
                horizontal = 'JustifyDistributed';
                break;
        }

        if (horizontal) {
            this._out.write(' ss:Horizontal="' + horizontal + '"');
        }

        if (style.alignment.indent > 0) { // 0 is default
            this._out.write(' ss:Indent="' + style.alignment.indent + '"');
        }

        var readingOrder = null;
        switch (style.alignment.readingOrder) {
            default:
            case HorizontalReadingOrder.Context: // Default
                break;
            case HorizontalReadingOrder.RightToLeft:
                readingOrder = 'RightToLeft';
                break;
            case HorizontalReadingOrder.LeftToRight:
                readingOrder = 'LeftToRight';
                break;
        }

        if (readingOrder) {
            this._out.write(' ss:ReadingOrder="' + readingOrder + '"');
        }

        if (style.alignment.rotate != 0) { // 0 is default
            this._out.write(' ss:Rotate="' + style.alignment.rotate + '"');
        }

        if (style.alignment.shrinkToFit) { // FALSE is default
            this._out.write(' ss:ShrinkToFit="1"');
        }

        var vertical = null;
        switch (style.alignment.vertical) {
            default:
            case VerticalAlignment.Automatic: // Default
                break;
            case VerticalAlignment.Top:
                vertical = 'Top';
                break;
            case VerticalAlignment.Bottom:
                vertical = 'Bottom';
                break;
            case VerticalAlignment.Center:
                vertical = 'Center';
                break;
            case VerticalAlignment.Justify:
                vertical = 'Justify';
                break;
            case VerticalAlignment.Distributed:
                vertical = 'Distributed';
                break;
            case VerticalAlignment.JustifyDistributed:
                vertical = 'JustifyDistributed';
                break;
        }

        if (vertical) {
            this._out.write(' ss:Vertical="' + vertical + '"');
        }

        if (style.alignment.verticalText) { // FALSE is default
            this._out.write(' ss:VerticalText="1"');
        }

        if (style.alignment.wrapText) { // FALSE is default
            this._out.write(' ss:WrapText="1"');
        }

        this._out.write('/>\n'); // Closing tag
    }

    if (typeof style.numberFormat === 'string') {
        if (style.numberFormat.length) {
            this._out.write('   <ss:NumberFormat ss:Format="' + style.numberFormat + '"/>\n', 'utf8');
        }
        else {
            this._out.write('   <ss:NumberFormat />\n');
        }
    }

    if (style.borders && style.borders.length > 0) {
        this._out.write('    <ss:Borders>'); // Opening tag

        for (var i = 0; i < style.borders.length; i++) {
            var border = style.borders[i];

            this._out.write('<ss:Border'); // Opening tag

            var position = null;
            switch (border.position) {
                default:
                case BorderPosition.Left:
                    position = 'Left';
                    break;
                case BorderPosition.Top:
                    position = 'Top';
                    break;
                case BorderPosition.Right:
                    position = 'Right';
                    break;
                case BorderPosition.Bottom:
                    position = 'Bottom';
                    break;
                case BorderPosition.DiagonalLeft:
                    position = 'DiagonalLeft';
                    break;
                case BorderPosition.DiagonalRight:
                    position = 'DiagonalRight';
                    break;
            }

            this._out.write(' ss:Position="' + position + '"'); // Required

            var borderColor = border.color ? Color(border.color) : null;
            if (borderColor && borderColor.alpha() > 0) {
                this._out.write(' ss:Color="' + colorForSpreadsheet(borderColor) + '"');
            }

            var lineStyle = null;
            switch (border.lineStyle) {
                default:
                case BorderLineStyle.None: // Default
                    break;
                case BorderLineStyle.Continuous:
                    lineStyle = 'Continuous';
                    break;
                case BorderLineStyle.Dash:
                    lineStyle = 'Dash';
                    break;
                case BorderLineStyle.Dot:
                    lineStyle = 'Dot';
                    break;
                case BorderLineStyle.DashDot:
                    lineStyle = 'DashDot';
                    break;
                case BorderLineStyle.DashDotDot:
                    lineStyle = 'DashDotDot';
                    break;
                case BorderLineStyle.SlantDashDot:
                    lineStyle = 'SlantDashDot';
                    break;
                case BorderLineStyle.Double:
                    lineStyle = 'Double';
                    break;
            }

            if (lineStyle) {
                this._out.write(' ss:LineStyle="' + lineStyle + '"');
            }

            if (border.weight > 0) { // 0 is default
                this._out.write(' ss:Weight="' + border.weight + '"');
            }

            this._out.write('/>\n'); // Closing tag
        }

        this._out.write('</ss:Borders>\n'); // Closing tag
    }

    if (style.interior) {
        this._out.write('   <ss:Interior'); // Opening tag

        var interiorColor = style.interior.color ? Color(style.interior.color) : null;
        if (interiorColor && interiorColor.alpha() > 0) {
            this._out.write(' ss:Color="' + colorForSpreadsheet(interiorColor) + '"');
        }

        var pattern = null;
        switch (style.interior.pattern) {
            default:
            case InteriorPattern.None: // Default
                break;
            case InteriorPattern.Solid:
                pattern = 'Solid';
                break;
            case InteriorPattern.Gray75:
                pattern = 'Gray75';
                break;
            case InteriorPattern.Gray50:
                pattern = 'Gray50';
                break;
            case InteriorPattern.Gray25:
                pattern = 'Gray25';
                break;
            case InteriorPattern.Gray125:
                pattern = 'Gray125';
                break;
            case InteriorPattern.Gray0625:
                pattern = 'Gray0625';
                break;
            case InteriorPattern.HorzStripe:
                pattern = 'HorzStripe';
                break;
            case InteriorPattern.VertStripe:
                pattern = 'VertStripe';
                break;
            case InteriorPattern.ReverseDiagStripe:
                pattern = 'ReverseDiagStripe';
                break;
            case InteriorPattern.DiagCross:
                pattern = 'DiagCross';
                break;
            case InteriorPattern.ThickDiagCross:
                pattern = 'ThickDiagCross';
                break;
            case InteriorPattern.ThinHorzStripe:
                pattern = 'ThinHorzStripe';
                break;
            case InteriorPattern.ThinVertStripe:
                pattern = 'ThinVertStripe';
                break;
            case InteriorPattern.ThinReverseDiagStripe:
                pattern = 'ThinReverseDiagStripe';
                break;
            case InteriorPattern.ThinDiagStripe:
                pattern = 'ThinDiagStripe';
                break;
            case InteriorPattern.ThinHorzCross:
                pattern = 'ThinHorzCross';
                break;
            case InteriorPattern.ThinDiagCross:
                pattern = 'ThinDiagCross';
                break;
        }

        if (pattern) {
            this._out.write(' ss:Pattern="' + pattern + '"');
        }

        var interiorPatternColor = style.interior.patternColor ? Color(style.interior.patternColor) : null;
        if (interiorPatternColor && interiorPatternColor.alpha() > 0) {
            this._out.write(' ss:PatternColor="' + colorForSpreadsheet(interiorPatternColor) + '"');
        }

        this._out.write('/>\n'); // Closing tag
    }

    if (style.font) {
        this._out.write('   <ss:Font'); // Opening tag

        if (style.font.bold) { // FALSE is default
            this._out.write(' ss:Bold="1"');
        }

        var fontColor = style.font.color ? Color(style.font.color) : null;
        if (fontColor && fontColor.alpha() > 0) {
            this._out.write(' ss:Color="' + colorForSpreadsheet(fontColor) + '"');
        }

        if (style.font.fontName) {
            this._out.write(' ss:FontName="' + style.font.fontName + '"', 'utf8');
        }

        if (style.font.italic) { // FALSE is default
            this._out.write(' ss:Italic="1"');
        }

        if (style.font.outline) { // FALSE is default
            this._out.write(' ss:Outline="1"');
        }

        if (style.font.shadow) { // FALSE is default
            this._out.write(' ss:Shadow="1"');
        }

        if (style.font.size != null && style.font.size != 10) { // 10 is default
            this._out.write(' ss:Size="' + style.font.size + '"');
        }

        if (style.font.strikeThrough) { // FALSE is default
            this._out.write(' ss:StrikeThrough="1"');
        }

        var underline = null;
        switch (style.font.underline) {
            default:
            case FontUnderline.None: // Default
                break;
            case FontUnderline.Single:
                underline = 'Single';
                break;
            case FontUnderline.Double:
                underline = 'Double';
                break;
            case FontUnderline.SingleAccounting:
                underline = 'SingleAccounting';
                break;
            case FontUnderline.DoubleAccounting:
                underline = 'DoubleAccounting';
                break;
        }

        if (underline != null) {
            this._out.write(' ss:Underline="' + underline + '"');
        }

        var verticalAlign = null;
        switch (style.font.verticalAlign) {
            default:
            case FontVerticalAlign.None: // Default
                break;
            case FontVerticalAlign.Subscript:
                verticalAlign = 'Subscript';
                break;
            case FontVerticalAlign.Superscript:
                verticalAlign = 'Superscript';
                break;
        }

        if (verticalAlign != null) {
            this._out.write(' ss:VerticalAlign="' + verticalAlign + '"');
        }

        if (style.font.charset > 0) { // 0 is default
            this._out.write(' ss:CharSet="' + style.font.charset + '"');
        }

        var family = null;
        switch (style.font.family) {
            default:
            case FontFamily.Automatic: // Default
                break;
            case FontFamily.Decorative:
                family = 'Decorative';
                break;
            case FontFamily.Modern:
                family = 'Modern';
                break;
            case FontFamily.Roman:
                family = 'Roman';
                break;
            case FontFamily.Script:
                family = 'Script';
                break;
            case FontFamily.Swiss:
                family = 'Swiss';
                break;
        }

        if (family != null) {
            this._out.write(' ss:Family="' + family + '"');
        }

        this._out.write('/>\n'); // Closing tag
    }

    this._out.write('  </Style>\n');

    return this;
};

/**
 * @module
 * @type {XmlSpreadsheetWriter}
 * */
module.exports = XmlSpreadsheetWriter;