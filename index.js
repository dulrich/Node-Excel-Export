var fs = require('fs');
var temp = require('temp').track();
var path = require('path');
var JSZip = require('jszip');
var async = require('async');
Date.prototype.getJulian = function () {
    return Math.floor(this / 86400000 - this.getTimezoneOffset() / 1440 + 2440587.5);
};
Date.prototype.oaDate = function () {
    return (this - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};

var fragments = require('./fragments');

var templateXLSX = new Buffer(fragments.templateXLSX,'base64');

var sheetFront = fragments.sheetFront;
var sheetBack  = fragments.sheetBack;

var sharedStringsFront = fragments.sharedFront;
var sharedStringsBack  = fragments.sharedBack;

exports.execute = function(config, callback) {
    var cols = config.cols,
        data = config.rows,
        dataRows = [],
        colsLength = cols.length,
        p,
        files = [],
        styleIndex,
        k = 0,
        cn = 1,
        dirPath,
        sharedString = {
            values: [],
            index: {},
            converted: []
        },
        sheet,
        sheetPos = 0;

    var write = function(str, callback) {
        var buf = new Buffer(str)
        var off = 0;
        var written = 0;
        return async.whilst(function () {
            return written < buf.length;
        }, function (callback) {
            fs.write(sheet, buf, off, buf.length - off, sheetPos, function (err, w) {
                if (err) {
                    return callback(err);
                }
                written += w;
                off += w;
                sheetPos += w;
                return callback();
            });
        }, callback);
    };
    return async.waterfall([
        function (callback) {
            return temp.mkdir('xlsx', function (err, dir) {
                if (err) {
                    return callback(err);
                }
                dirPath = dir;
                return callback();
            });
        },
        function (callback) {
            return fs.mkdir(path.join(dirPath, 'xl'), callback);
        },
        function (callback) {
            return fs.mkdir(path.join(dirPath, 'xl', 'worksheets'), callback);
        },
        function (callback) {
            return async.parallel([
                function (callback) {
                    return fs.writeFile(path.join(dirPath, 'data.zip'), templateXLSX, callback);
                },
                function (callback) {
                    if (!config.stylesXmlFile) {
                        return callback();
                    }
                    p = config.stylesXmlFile || __dirname + '/styles.xml';
                    return fs.readFile(p, 'utf8', function (err, styles) {
                        if (err) {
                            return callback(err);
                        }
                        p = path.join(dirPath, 'xl', 'styles.xml');
                        files.push(p);
                        return fs.writeFile(p, styles, callback);
                    });
                }
            ], function (err) {
                return callback(err);
            });
        },
        function (callback) {
            p = path.join(dirPath, 'xl', 'worksheets', 'sheet.xml');
            files.push(p);
            return fs.open(p, 'a+', function (err, fd) {
                if (err) {
                    return callback(err);
                }
                sheet = fd;
                return callback();
            });
        },
        function (callback) {
            return write(sheetFront, callback);
        },
        function (callback) {
            return async.eachSeries(cols, function (col, callback) {
                var colStyleIndex = col.styleIndex || 0;
                var res = '<x:col min="' + cn + '" max="' + cn + '" width="' + (col.width ? col.width : 10) + '" customWidth="1" style="' + colStyleIndex + '"/>';
                cn++;
                return write(res, callback);
            }, callback);
        },
        function (callback) {
            return write('</cols><x:sheetData>', callback);
        },
        function (callback) {
            return write('<x:row r="1" spans="1:' + colsLength + '">', callback);
        },
        function (callback) {
            return async.eachSeries(cols, function (col, callback) {
                var colStyleIndex = col.captionStyleIndex || 0;
                var res = addStringCol(getColumnLetter(k + 1) + 1, col.caption, colStyleIndex, sharedString);

                return write(res, callback);
            }, callback);
        },
        function (callback) {
            return write('</x:row>', callback);
        },
        function (callback) {
            var j, r, cellData, currRow, cellType;

            function beforeCellWrite(row,cellData,eOpt) {
                var type;

                type = typeof cellData;
                eOpt.cellType = 'string';

                if (type === 'string') {
                    return cellData;
                }
                else if (type === 'number') {
                    eOpt.cellType = 'number';
                    return cellData;
                }
                else if (type === 'boolean') {
                    eOpt.cellType = 'bool';
                    return cellData.toString();
                }
                else if (cellData instanceof Date) {
                    if (!cellData) return '';

                    eOpt.cellType = 'date';
                    return moment(cellData).toDate().oaDate();
                }
                else if (cellData instanceof Array) {
                    return cellData.join(',');
                }
                else if (type === 'object') {
                    return JSON.stringify(cellData);
                }
                else {
                    return '';
                }
            }

            dataRows = data.map(function(r,i) {
                currRow = i+2;

                var row = '<x:row r="' + currRow +'" spans="1:'+ colsLength + '">';

                for (j=0; j < colsLength; j++) {
                    styleIndex = null;
                    cellData = r[j];
                    cellType = cols[j].type;

                    cols[j].beforeCellWrite = (typeof cols[j].beforeCellWrite === 'function')
                        ? cols[j].beforeCellWrite
                        : beforeCellWrite;

                    var e = {
                        rowNum: currRow,
                        styleIndex: null,
                        cellType: cellType
                    };

                    cellData = cols[j].beforeCellWrite(r, cellData, e);
                    styleIndex = e.styleIndex || styleIndex;
                    cellType = e.cellType;
                    e = undefined;

                    switch (cellType) {
                        case 'number':
                            row += addNumberCol(getColumnLetter(j+1)+currRow, cellData, styleIndex);
                            break;
                        case 'date':
                            row += addDateCol(getColumnLetter(j+1)+currRow, cellData, styleIndex);
                            break;
                        case 'bool':
                            row += addBoolCol(getColumnLetter(j+1)+currRow, cellData, styleIndex);
                            break;
                        default:
                            row += addStringCol(getColumnLetter(j+1)+currRow, cellData, styleIndex, sharedString);
                            break;
                    }
                }

                row += '</x:row>';

                return row;
            });

            callback();
        },
        function(callback) {
            async.eachSeries(dataRows,function(row,callback) {
                return write(row,callback);
            },callback);
        },
        function (callback) {
            return write(sheetBack, callback);
        },
        function (callback) {
            return fs.close(sheet, callback);
        },
        function (callback) {
            if (sharedString.values.length === 0) {
                return callback();
            }
            sharedStringsFront = sharedStringsFront.replace(/\$count/g, sharedString.values.length);
            p = path.join(dirPath, 'xl', 'sharedStrings.xml');
            files.push(p);
            return fs.writeFile(p, sharedStringsFront + sharedString.converted.join("") + sharedStringsBack, callback);
        }
    ], function (err) {
        if (err) {
            return callback(err);
        }
        var prev = fs.readFileSync(path.join(dirPath, 'data.zip'));
        var zip = new JSZip(prev);
        files.forEach(function (file) {
            var relative = path.relative(dirPath, file);
            zip.file(relative, fs.readFileSync(file));
        });
        var data = zip.generate({
            mimeType: 'application/zip',
            type: 'nodebuffer'
        });
        temp.cleanup();
        return callback(null, data);
    });
};
var startTag = function (obj, tagName, closed) {
    var result = '<' + tagName, p;
    for (p in obj) {
        result += ' ' + p + '=' + obj[p];
    }
    if (!closed) {
        result += '>';
    } else {
        result += '/>';
    }
    return result;
};
var endTag = function (tagName) {
    return '</' + tagName + '>';
};
var addNumberCol = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return '';
    } else {
        return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
    }
};
var addDateCol = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 1;
    if (value === null) {
        return '';
    } else {
        return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
    }
};
var addBoolCol = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return '';
    }
    if (value) {
        value = 1;
    } else {
        value = 0;
    }
    return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="b"><x:v>' + value + '</x:v></x:c>';
};
var addStringCol = function (cellRef, value, styleIndex, sharedString) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return [
            '',
            ''
        ];
    }
    if (typeof value === 'string') {
        value = value.replace(/&/g, '&amp;').replace(/'/g, '&apos;').replace(/>/g, '&gt;').replace(/</g, '&lt;');
    }

    var i = sharedString.index[value] || -1;
    if ( i < 0) {
        i = sharedString.values.push(value) - 1;
        sharedString.index[value] = i;
        sharedString.converted.push("<x:si><x:t>" + value + "</x:t></x:si>");
    }

    return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="s"><x:v>' + i + '</x:v></x:c>';
}

var getColumnLetter = function(col) {
    if (col <= 0) {
        throw 'col must be more than 0';
    }
    var array = [];
    while (col > 0) {
        var remainder = col % 26;
        col /= 26;
        col = Math.floor(col);
        if (remainder === 0) {
            remainder = 26;
            col--;
        }
        array.push(64 + remainder);
    }
    return String.fromCharCode.apply(null, array.reverse());
};
