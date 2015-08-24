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
    XmlSpreadsheetWriter: /** @type {XmlSpreadsheetWriter} */ require('./xml/writer'),

    /**
     * @public
     * @type {CsvSpreadsheetWriter}
     * */
    CsvSpreadsheetWriter: /** @type {CsvSpreadsheetWriter} */ require('./csv/writer')

};