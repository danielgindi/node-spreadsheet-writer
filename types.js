
/**
 * @enum {string}
 */
var VerticalAlignment = {
    Automatic: 'Automatic',
    Top: 'Top',
    Center: 'Center',
    Bottom: 'Bottom',
    Justify: 'Justify',
    Distributed: 'Distributed',
    JustifyDistributed: 'JustifyDistributed'
};

/**
 * @enum {string}
 */
var BorderLineStyle = {
    None: 'None',
    Continuous: 'Continuous',
    Dash: 'Dash',
    Dot: 'Dot',
    DashDot: 'DashDot',
    DashDotDot: 'DashDotDot',
    SlantDashDot: 'SlantDashDot',
    Double: 'Double'
};

/**
 * @enum {string}
 */
var BorderPosition = {
    Left: 'Left',
    Top: 'Top',
    Right: 'Right',
    Bottom: 'Bottom',
    DiagonalLeft: 'DiagonalLeft',
    DiagonalRight: 'DiagonalRight'
};

/**
 * @enum {string}
 */
var FontFamily = {
    Automatic: 'Automatic',
    Decorative: 'Decorative',
    Modern: 'Modern',
    Roman: 'Roman',
    Script: 'Script',
    Swiss: 'Swiss'
};

/**
 * @enum {string}
 */
var FontUnderline = {
    None: 'None',
    Single: 'Single',
    Double: 'Double',
    SingleAccounting: 'SingleAccounting',
    DoubleAccounting: 'DoubleAccounting'
};

/**
 * @enum {string}
 */
var FontVerticalAlign = {
    None: 'None',
    Subscript: 'Subscript',
    Superscript: 'Superscript'
};

/**
 * @enum {string}
 */
var HorizontalAlignment = {
    Automatic: 'Automatic',
    Left: 'Left',
    Center: 'Center',
    Right: 'Right',
    Fill: 'Fill',
    Justify: 'Justify',
    CenterAcrossSelection: 'CenterAcrossSelection',
    Distributed: 'Distributed',
    JustifyDistributed: 'JustifyDistributed'
};

/**
 * @enum {string}
 */
var HorizontalReadingOrder = {
    Context: 'Context',
    RightToLeft: 'RightToLeft',
    LeftToRight: 'LeftToRight'
};

/**
 * @enum {string}
 */
var InteriorPattern = {
    None: 'None',
    Solid: 'Solid',
    Gray75: 'Gray75',
    Gray50: 'Gray50',
    Gray25: 'Gray25',
    Gray125: 'Gray125',
    Gray0625: 'Gray0625',
    HorzStripe: 'HorzStripe',
    VertStripe: 'VertStripe',
    ReverseDiagStripe: 'ReverseDiagStripe',
    DiagStripe: 'DiagStripe',
    DiagCross: 'DiagCross',
    ThickDiagCross: 'ThickDiagCross',
    ThinHorzStripe: 'ThinHorzStripe',
    ThinVertStripe: 'ThinVertStripe',
    ThinReverseDiagStripe: 'ThinReverseDiagStripe',
    ThinDiagStripe: 'ThinDiagStripe',
    ThinHorzCross: 'ThinHorzCross',
    ThinDiagCross: 'ThinDiagCross'
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
    Number: 'Number',
    DateTime: 'DateTime',
    Boolean: 'Boolean',
    String: 'String',
    Error: 'Error'
};

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