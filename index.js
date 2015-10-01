/**
 * @module
 */

module.exports = {

    /** @public */
    Types: require('./types'),

    /** @public */
    Utils: require('./utils'),

    /**
     * @public
     * @type {function(new:XmlSpreadsheetWriter, *)}
     * */
    XmlSpreadsheetWriter: require('./xml'),

    /**
     * @public
     * @type {function(new:CsvSpreadsheetWriter, *)}
     * */
    CsvSpreadsheetWriter: require('./csv')

};