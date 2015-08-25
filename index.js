/**
 * @module
 */

module.exports = {

    /** @public */
    Types: require('./types'),

    /**
     * @public
     * @type {XmlSpreadsheetWriter}
     * */
    XmlSpreadsheetWriter: /** @type {XmlSpreadsheetWriter} */ require('./xml'),

    /**
     * @public
     * @type {CsvSpreadsheetWriter}
     * */
    CsvSpreadsheetWriter: /** @type {CsvSpreadsheetWriter} */ require('./csv')

};