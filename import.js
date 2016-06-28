'use strict';

let outputDir = 'output'; // default

const path = require('path');
const xslt4node = require('xslt4node');
const extract = require('extract-zip');
const fs = require('fs-extra');
const DOMParser = require('xmldom').DOMParser;
const XMLSerializer = require('xmldom').XMLSerializer;
const async = require('async');
const _ = require('lodash');
const Q = require('q');
const mv = require('mv');
const fileExists = require('file-exists');
let config = {};

/**
 * Constructor
 * @param {object} options - config options
 * @returns {Convert}
 * @constructor
 */
function Convert(options) {
    if (!(this instanceof Convert)) {
        return new Convert(options);
    }
    this.init();
    config = options;
}

/**
 * Set up object: read in XSL for transforms
 */
Convert.prototype.init = function () {
    fs.readFile(process.cwd() + "/wordtoxml.xsl", (err, data) => {
        if (err) {
            return console.error(err);
        }
        this.XML2HTML = new DOMParser().parseFromString(data.toString());
    });
    this.WordXML = null;
    console.log('Initialized.')
};

/**
 * Main import function
 * @param {string } sourcePath - path to input docx file.
 * @param {string } outDir - output folder
 */
Convert.prototype.import = function (sourcePath, outDir) {

    if (outDir) {
        outputDir = outDir;
    }

    this.WordXML = null;

    let sourceFile = {
        path: sourcePath,
        name: path.basename(sourcePath),
        dir: path.dirname(sourcePath) + '',
        extname: path.extname(sourcePath),
        basename: path.basename(sourcePath, path.extname(sourcePath))
    };
    sourceFile.outDir = outputDir + '/' + sourceFile.basename;

    // unzip the docx file, then get the xml files and transform
    let _extract = function () {
        extract(sourceFile.path, {dir: sourceFile.outDir}, function (err) {
            if (err) {
                return console.error(err);
            }
            _getXMLAndTransform();
        });
    };

    /**
     * Need to read several XML files:
     *   the main document.xml,
     *   the rels file (which contains links to the images),
     *   the core.xml file which has properties like title etc.
     * Append one to the other, and then we can call transform.
     *
     * The next two functions perform these tasks
     * @param {string} xmlPath to xml file
     * @param {function} callback (so we can use with async library, see _getXML)
     * @private
     */
    let _addXML = (xmlPath, callback) => {
        var filePath = sourceFile.outDir + '/' + xmlPath;
        fs.readFile(filePath, (err, data) => {
            if (err) {
                if ('ENOENT' !== err.code) {
                    console.error("addXML:", err);
                }
                if ('ENOENT' === err.code) {
                    console.log('File not found in add XML:', filePath);
                }
                return callback();
            }
            let xml = new DOMParser().parseFromString(data.toString());
            if (!this.WordXML) {
                this.WordXML = xml;
            } else {
                this.WordXML.documentElement.appendChild(
                    this.WordXML.importNode(xml.documentElement, true)
                );
            }
            callback();
        });
    };

    /**
     * Put together the separate xml files unzipped from the docx file,
     * then transform them with XSL.
     * @private
     */
    let _getXMLAndTransform = function () {
        let paths = [
            'word/document.xml',
            'word/_rels/document.xml.rels',
            'docProps/core.xml',
            'word/styles.xml',
            'word/numbering.xml'
        ];
        let funcs = [];
        paths.forEach((path) => {
            funcs.push((callback) => {
                _addXML(path, callback)
            })
        });
        funcs.push(() => {
            _writeWordXML().then(()=> _transform());
        });
        async.series(funcs);
    };

    /**
     * Once we have all of the doc xml files loaded into a DOM object,
     * we will apply the XSL to get the intermediate XML which we can
     * easily transform into JSON and other formats
     * @private
     */
    let _transform = () => {

        const xslt = new XMLSerializer().serializeToString(this.XML2HTML);
        const xml = new XMLSerializer().serializeToString(this.WordXML);
        const config = {
            xslt: xslt,
            source: xml,
            result: String,
            props: {
                indent: 'yes'
            }
        };
        xslt4node.transform(config, (err, result) => {
            _writeXML(result, '/word/output_document.xml')
                .then(() => _createJsonOutput(result))
                .then(() => _cleanup());
        });

    };

    /**
     * Create Json output from xml.
     *
     * @param {string} xmlString - input xml as String
     * @private
     */
    let _createJsonOutput = function (xmlString) {

        let doc = new DOMParser().parseFromString(xmlString);

        let wordJson = {};
        let items = wordJson.items = [];

        let tocEl = doc.getElementsByTagName("toc")[0];
        if (tocEl) {
            let toc = {};
            let heading = tocEl.getElementsByTagName("heading")[0];
            if (heading) {
                toc.heading = heading.textContent
            }
            let linkEls = tocEl.getElementsByTagName("links")[0];
            if (linkEls) {
                linkEls = linkEls.getElementsByTagName("link");
                toc.links = [];
                for (let l = 0; l < linkEls.length; l++) {
                    let linkEl = linkEls[l];
                    let link = {};
                    _attachAttrs(link, linkEl);
                    toc.links.push(link);
                }
            }
            wordJson.toc = toc;
        }
        // build JSON for the items
        let itemEls = doc.getElementsByTagName("item");
        _.forEach(itemEls, (itemEl) => {
            let item = {};
            _attachAttrs(item, itemEl);
            // get the item attributes, put in JSON
            let elType = itemEl.getAttribute("type");
            // convert the item content to JSON
            let contentEl = itemEl.getElementsByTagName("content")[0];
            if (contentEl) {
                let content = new XMLSerializer().serializeToString(contentEl);
                content = content.replace(/^<content>/, '').replace(/<\/content>$/, '').replace(/\n/g, '').replace(/<content\/>/, '');
                item.content = content;
                if (elType === "table") {
                    let rowEls = contentEl.getElementsByTagName("tr");
                    let rows = [];
                    _.forEach(rowEls, (rowEl) => {
                        let row = [];
                        let colEls = rowEl.getElementsByTagName("td");
                        if (colEls) {
                            _.forEach(colEls, (colEl) => {
                                row.push(colEl.textContent);
                            });
                        }
                        rows.push(row);
                    });
                    item.rows = rows;
                }
                if (elType == "list") {
                    let list = getJsonFromDomList(null, contentEl);
                    item.list = list[0];
                }
            }
            if (item.content) {
                items.push(item);
            }
        });
        let headEls = doc.getElementsByTagName("head")[0].childNodes;
        if (!headEls) {
            headEls = []
        }
        _.forEach(headEls, (el) => {
            if (el.nodeType === 1) {
                wordJson[el.tagName] = el.textContent;
            }
        });
        return _writeWordJson(wordJson);

    };

    // util to take attributes from DOM El and put on object as keys
    let _attachAttrs = function (item, el) {
        _.forEach(el.attributes, (attr) => {
            let v = attr.textContent;
            if (v === 'true') {
                v = true
            }
            if (v === 'false') {
                v = false
            }
            item[attr.name] = v;
        });
    };

    let _writeWordXML = () => {
        return _writeXML(this.WordXML, '/word/imported_document.xml');
    };

    let _writeXML = (xml, filePath) => {
        let deferred = Q.defer();
        let xmlString = new XMLSerializer().serializeToString(xml);
        fs.writeFile(sourceFile.outDir + filePath, xmlString, (err) => {
            if (err) {
                console.log(err);
                return deferred.reject(err);
            }
            return deferred.resolve();
        });
        return deferred.promise;
    };

    let _writeWordJson = (wordJson) => {
        var deferred = Q.defer();
        var outPath = sourceFile.outDir + '/' + sourceFile.basename + ".json";
        fs.writeFile(outPath, JSON.stringify(wordJson), (err) => {
            if (err) {
                console.log(err);
                return deferred.reject(err);
            }
            console.log('Conversion finished.');
            return deferred.resolve();
        });
        return deferred.promise;
    };

    let _cleanup = ()=> {
        try {
            fs.copySync(sourceFile.outDir + '/word/media', sourceFile.outDir + '/media');
        } catch(e){
            // no media files
        }
        if (config.cleanup) {
            fs.removeSync(sourceFile.outDir + '/word');
            fs.removeSync(sourceFile.outDir + '/_rels');
            fs.removeSync(sourceFile.outDir + '/docProps');
            fs.removeSync(sourceFile.outDir + '/*.xml');
        }
        process.exit(0);
    };

    // start the import process
    _extract(path);

};

/**
 * Utility function, takes an array and an element, gets all
 * child data from list item nodes. Returns nested arrays
 * corresponding to the DOM subtree.
 * @param list
 * @param el
 * @returns {*}
 */
function getJsonFromDomList(list, el) {
    if (el.tagName === 'li') {
        return list.push(el.textContent);
    }
    let children = el.childNodes;
    if (!children) {
        return;
    }
    let childList = [];
    if (list) {
        list.push(childList);
    } else {
        list = childList;
    }
    _.forEach(children, (child) => {
        if (child.nodeType !== 3) {
            getJsonFromDomList(childList, child);
        }
    });
    return list;
}

module.exports = Convert;