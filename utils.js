//var mkdirp = require('mkdirp');

var Utils = {

    toOADate: (function () {
        /** @const */ var utc18991230 = Date.UTC(1899, 11, 31);
        /** @const */ var msPerDay = 24 * 60 * 60 * 1000;

        return function (date) {
            if (date instanceof Date) {
                date = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
            }
            return (date - utc18991230) / msPerDay;
        };

    })()

    , toXMLDateTime: function (date) {
        var dateTime = date.getFullYear() + '-';

        var i = date.getMonth() + 1;
        dateTime += (i < 10 ? '0' + i : i) + '-';

        i = date.getDate();
        dateTime += (i < 10 ? '0' + i : i) + 'T';

        i = date.getHours();
        dateTime += (i < 10 ? '0' + i : i) + ':';

        i = date.getMinutes();
        dateTime += (i < 10 ? '0' + i : i) + ':';

        i = date.getSeconds();
        dateTime += (i < 10 ? '0' + i : i) + '.';

        i = date.getMilliseconds();
        dateTime += (i < 10 ? '00' + i : (i < 100 ? '0' + i : i));

        return dateTime;
    }

    , stripHtml: (function () {

        var AllHtmlEntities = new (require('html-entities').AllHtmlEntities);

        return function (html) {

            // Remove BR tags and replace them with newlines
            html = html.replace(/<br[^>]*>/gi, '\n');

            // Strips the <script> tags from the Html
            html = html.replace(/<script[^>.]*>[\s\S]*?<\/script>/ig, ' ');

            // Strips the <style> tags from the Html
            html = html.replace(/<style[^>.]*>[\s\S]*?<\/style>/ig, ' ');

            // Strips the <!--comment--> tags from the Html
            html = html.replace(/<!(?:--[\s\S]*?--\s*)?>/ig, ' ');

            // Strips inline tags
            html = html.replace(/<\/?(a|b|big|i|small|tt|abbr|acronym|dfn|em|strong|samp|var|a|bdo|span)[^>]*>/ig, '');

            // Strips block tags
            html = html.replace(/<(div|p)[^>]*>/ig, '\n').replace(/<\/p>/ig, '\n').replace(/<\/div>/ig, '');

            // Strips the HTML tags from the Html
            html = html.replace(/<(.|\n)+?>/ig, ' ');

            // Decode all html entities
            html = AllHtmlEntities.decode(html);

            return html;
        };

    })()

    /*, mkdirpMulti: function (dirs, mode, callback) {
        if (typeof mode === 'function') {
            callback = /* * @type {function} * / mode;
            mode = undefined;
        }

        if (!Array.isArray(dirs)) {
            dirs = dirs ? [dirs] : [];
        }

        var index = 0;
        var next = function () {
            if (index >= dirs.length) {
                callback && callback();
                return;
            }

            mkdirp(dirs[index++], mode, function (err) {

                if (err) {
                    callback && callback();
                    return;
                }

                next();

            });
        };

        next();

        return this;
    }*/

};

module.exports = Utils;