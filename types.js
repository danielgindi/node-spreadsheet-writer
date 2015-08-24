
/**
 * @enum {string}
 */
var VerticalAlignment = {
    /** @const */ Automatic: 'Automatic',
    /** @const */ Top: 'Top',
    /** @const */ Center: 'Center',
    /** @const */ Bottom: 'Bottom',
    /** @const */ Justify: 'Justify',
    /** @const */ Distributed: 'Distributed',
    /** @const */ JustifyDistributed: 'JustifyDistributed'
};

/**
 * @enum {string}
 */
var BorderLineStyle = {
    /** @const */ None: 'None',
    /** @const */ Continuous: 'Continuous',
    /** @const */ Dash: 'Dash',
    /** @const */ Dot: 'Dot',
    /** @const */ DashDot: 'DashDot',
    /** @const */ DashDotDot: 'DashDotDot',
    /** @const */ SlantDashDot: 'SlantDashDot',
    /** @const */ Double: 'Double'
};

/**
 * @enum {string}
 */
var BorderPosition = {
    /** @const */ Left: 'Left',
    /** @const */ Top: 'Top',
    /** @const */ Right: 'Right',
    /** @const */ Bottom: 'Bottom',
    /** @const */ DiagonalLeft: 'DiagonalLeft',
    /** @const */ DiagonalRight: 'DiagonalRight'
};

/**
 * @enum {string}
 */
var FontFamily = {
    /** @const */ Automatic: 'Automatic',
    /** @const */ Decorative: 'Decorative',
    /** @const */ Modern: 'Modern',
    /** @const */ Roman: 'Roman',
    /** @const */ Script: 'Script',
    /** @const */ Swiss: 'Swiss'
};

/**
 * @enum {string}
 */
var FontUnderline = {
    /** @const */ None: 'None',
    /** @const */ Single: 'Single',
    /** @const */ Double: 'Double',
    /** @const */ SingleAccounting: 'SingleAccounting',
    /** @const */ DoubleAccounting: 'DoubleAccounting'
};

/**
 * @enum {string}
 */
var FontVerticalAlign = {
    /** @const */ None: 'None',
    /** @const */ Subscript: 'Subscript',
    /** @const */ Superscript: 'Superscript'
};

/**
 * @enum {string}
 */
var HorizontalAlignment = {
    /** @const */ Automatic: 'Automatic',
    /** @const */ Left: 'Left',
    /** @const */ Center: 'Center',
    /** @const */ Right: 'Right',
    /** @const */ Fill: 'Fill',
    /** @const */ Justify: 'Justify',
    /** @const */ CenterAcrossSelection: 'CenterAcrossSelection',
    /** @const */ Distributed: 'Distributed',
    /** @const */ JustifyDistributed: 'JustifyDistributed'
};

/**
 * @enum {string}
 */
var HorizontalReadingOrder = {
    /** @const */ Context: 'Context',
    /** @const */ RightToLeft: 'RightToLeft',
    /** @const */ LeftToRight: 'LeftToRight'
};

/**
 * @enum {string}
 */
var InteriorPattern = {
    /** @const */ None: 'None',
    /** @const */ Solid: 'Solid',
    /** @const */ Gray75: 'Gray75',
    /** @const */ Gray50: 'Gray50',
    /** @const */ Gray25: 'Gray25',
    /** @const */ Gray125: 'Gray125',
    /** @const */ Gray0625: 'Gray0625',
    /** @const */ HorzStripe: 'HorzStripe',
    /** @const */ VertStripe: 'VertStripe',
    /** @const */ ReverseDiagStripe: 'ReverseDiagStripe',
    /** @const */ DiagStripe: 'DiagStripe',
    /** @const */ DiagCross: 'DiagCross',
    /** @const */ ThickDiagCross: 'ThickDiagCross',
    /** @const */ ThinHorzStripe: 'ThinHorzStripe',
    /** @const */ ThinVertStripe: 'ThinVertStripe',
    /** @const */ ThinReverseDiagStripe: 'ThinReverseDiagStripe',
    /** @const */ ThinDiagStripe: 'ThinDiagStripe',
    /** @const */ ThinHorzCross: 'ThinHorzCross',
    /** @const */ ThinDiagCross: 'ThinDiagCross'
};

/**
 * @typedef {{
 * horizontal: HorizontalAlignment = Automatic,
 * vertical: VerticalAlignment = Automatic,
 * indent: int = 0,
 * readingOrder: HorizontalReadingOrder = Context,
 * rotate: Number = 0,
 * shrinkToFit: boolean = false,
 * verticalText: boolean = false,
 * wrapText: boolean = false
 * }} Alignment
 */

/**
 * @typedef {{
 * position: BorderPosition,
 * color: String,
 * lineStyle: BorderLineStyle = None,
 * weight: Number= 0
 * }} Border
 */

/**
 * @typedef {{
 * bold: boolean = false,
 * color: String?,
 * fontName: String?,
 * italic: boolean = false,
 * outline: boolean = false,
 * shadow: boolean = false,
 * size: Number = 0,
 * strikeThrough: boolean = false,
 * underline: FontUnderline = None,
 * verticalAlign: FontVerticalAlign = None,
 * charset: int = 0,
 * family: FontFamily = Automatic
 * }} Font
 */

/**
 * @typedef {{
 * color: String?,
 * pattern: InteriorPattern = None,
 * patternColor: String?
 * }} Interior
 */

/**
 * @typedef {{
 * parentStyleIndex: int?,
 * numberFormat: String?,
 * alignment: Alignment = None,
 * borders: Array.<Border>?,
 * interior: Interior?,
 * font: Font?
 * }} SpreadsheetCellStyle
 */

var NumberFormats = {
    /** @const */ Automatic:  '',
    /** @const */ General: 'General',
    /** @const */ GeneralNumber: 'General Number',
    /** @const */ GeneralDate: 'General Date',
    /** @const */ LongDate: 'Long Date',
    /** @const */ MediumDate: 'Medium Date',
    /** @const */ ShortDate: 'Short Date',
    /** @const */ LongTime: 'Long Time',
    /** @const */ MediumTime: 'Medium Time',
    /** @const */ ShortTime: 'Short Time',
    /** @const */ Currency: 'Currency',
    /** @const */ EuroCurrency: 'Euro Currency',
    /** @const */ Fixed: 'Fixed',
    /** @const */ Standard: 'Standard',
    /** @const */ Percent: 'Percent',
    /** @const */ Scientific: 'Scientific',
    /** @const */ YesNo: 'Yes/No',
    /** @const */ TrueFalse: 'True/False',
    /** @const */ OnOff: 'On/Off',
    /** @const */ Number0: '0',
    /** @const */ Number0_00: '0.00'
};

/**
 * @enum {string}
 */
var CellType = {
    /** @const */ Number: 'Number',
    /** @const */ DateTime: 'DateTime',
    /** @const */ Boolean: 'Boolean',
    /** @const */ String: 'String',
    /** @const */ Error: 'Error'
};

/**
 * @enum {string}
 */
var CellFormulaPlaceholder = {
    /** @const */ Null: '#NULL!',
    /** @const */ DivisionByZero: '#DIV/0!',
    /** @const */ InvalidValue: '#VALUE!',
    /** @const */ InvalidReference: '#REF!',
    /** @const */ InvalidNameReference: '#NAME?',
    /** @const */ NotANumber: '#NUM!',
    /** @const */ NotAvailable: '#N/A',
    /** @const */ CircularReference: '#CIRC!'
};

var Types = { };

/** @public */ Types.VerticalAlignment = VerticalAlignment;
/** @public */ Types.BorderLineStyle = BorderLineStyle;
/** @public */ Types.BorderPosition = BorderPosition;
/** @public */ Types.FontFamily = FontFamily;
/** @public */ Types.FontUnderline = FontUnderline;
/** @public */ Types.FontVerticalAlign = FontVerticalAlign;
/** @public */ Types.HorizontalAlignment = HorizontalAlignment;
/** @public */ Types.HorizontalReadingOrder = HorizontalReadingOrder;
/** @public */ Types.InteriorPattern = InteriorPattern;
/** @public */ Types.NumberFormats = NumberFormats;
/** @public */ Types.CellType = CellType;
/** @public */ Types.CellFormulaPlaceholder = CellFormulaPlaceholder;


module.exports = Types;