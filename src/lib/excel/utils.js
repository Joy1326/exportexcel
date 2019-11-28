/* eslint-disable no-console */
import colCache from 'exceljs/lib/utils/col-cache';
window.colCache = colCache;
export function isDOM(str) {
    return str instanceof HTMLElement;
}
export function isString(str) {
    return typeOf(str) === 'string';
}
export function isObject(str) {
    return typeOf(str) === 'object';
}
export function isArray(str) {
    return typeOf(str) === 'array';
}
export function isEmpty(val) {
    return val === undefined || val === null;
}
export function isFunction(fcn) {
    return typeOf(fcn) === 'function';
}
export function encodeAddress2(row, col) {
    return encodeAddress(row + 1, col + 1);
}
export function encodeAddress(row, col) {
    return colCache.encodeAddress(row, col);
}
export function decodeAddress(value) {
    return colCache.decodeAddress(value);
}
function typeOf(obj) {
    const toString = Object.prototype.toString;
    const map = {
        '[object Boolean]': 'boolean',
        '[object Number]': 'number',
        '[object String]': 'string',
        '[object Function]': 'function',
        '[object Array]': 'array',
        '[object Date]': 'date',
        '[object RegExp]': 'regExp',
        '[object Undefined]': 'undefined',
        '[object Null]': 'null',
        '[object Object]': 'object'
    };
    return map[toString.call(obj)];
}
// deepCopy
function deepCopy(data) {
    const t = typeOf(data);
    let o;

    if (t === 'array') {
        o = [];
    } else if (t === 'object') {
        o = {};
    } else {
        return data;
    }

    if (t === 'array') {
        for (let i = 0; i < data.length; i++) {
            o.push(deepCopy(data[i]));
        }
    } else if (t === 'object') {
        for (let i in data) {
            o[i] = deepCopy(data[i]);
        }
    }
    return o;
}

export function getAllColumns(cols, forTableHead = false) {
    const columns = deepCopy(cols);
    const result = [];
    columns.forEach((column) => {
        if (column.children) {
            if (forTableHead) { result.push(column); }
            result.push.apply(result, getAllColumns(column.children, forTableHead));
        } else {
            result.push(column);
        }
    });
    return result;
}
// export function getKeys(columns) {
//     let keys = [];
//     let allColumns = getAllColumns(columns);
//     for (let i = 0, len = allColumns.length; i < len; i++) {
//         let { key } = allColumns[i];
//         keys.push(key);
//     }
//     return keys;
// }

export function getKeys(columns) {
    let keys = [];
    let allColumns = getAllColumns(columns);
    for (let i = 0, len = allColumns.length; i < len; i++) {
        // let { key, fmt,cellStyle } = allColumns[i];
        let opt = allColumns[i];
        removeSomeFcn(opt);
        keys.push(opt);
        // keys.push({
        //     key,
        //     fmt,
        //     cellStyle
        // });
    }
    return keys;
}
function removeSomeFcn(opt) {
    for (let o in opt) {
        if (isFunction(opt[o])) {
            if (o === 'fmt' || o === 'cellStyle') {
                continue;
            }
            delete opt[o];
        }
    }
}

export function convertToRows(columns, { startRow = 0, startCol = 0 } = { startRow: 0, startCol: 0 }) {
    try {
        const originColumns = deepCopy(columns);
        let maxLevel = 1;
        const traverse = (column, parent) => {
            if (parent) {
                column.level = parent.level + 1;
                if (maxLevel < column.level) {
                    maxLevel = column.level;
                }
            }

            if (column.children) {
                let colSpan = 0;
                column.children.forEach((subColumn) => {
                    traverse(subColumn, column);
                    colSpan += subColumn.colSpan;
                });
                column.colSpan = colSpan;
            } else {
                column.colSpan = 1;
            }
        };

        originColumns.forEach((column) => {
            column.level = 1;
            traverse(column);
        });

        const rows = [];
        for (let i = 0; i < maxLevel; i++) {
            rows.push([]);
        }

        const allColumns = getAllColumns(originColumns, true);
        let merges = [];
        // let cell = [];
        let cellInfo = {};
        let INDEX = 0;
        allColumns.forEach((column) => {
            let level = column.level;
            if (!column.children) {
                column.rowSpan = maxLevel - level + 1;
            }
            let rowSpan = column.rowSpan && column.rowSpan > 1 ? column.rowSpan - 1 : 0;
            let colSpan = column.colSpan && column.colSpan > 1 ? column.colSpan - 1 : 0;
            if (column.rowSpan > 1 || column.colSpan > 1) {
                merges.push({
                    // s: {
                    //     r: level - 1,
                    //     c: INDEX
                    // },
                    // e: {
                    //     r: level - 1 + rowSpan,
                    //     c: INDEX + colSpan
                    // },
                    s: encodeAddress2(level - 1 + startRow, INDEX + startCol),
                    e: encodeAddress2(level - 1 + rowSpan + startRow, INDEX + colSpan + startCol)
                });
            }
            // cell.push({
            // col: {
            //     r: level - 1,
            //     c: INDEX,
            // },
            // key: column.key,
            // title: column.title,
            // colSpan: colSpan,
            // rowSpan: rowSpan

            // });
            cellInfo[encodeAddress2(level - 1 + startRow, INDEX + startCol)] = column.title;
            if (column.children) {
                column.rowSpan = 1;
                INDEX--;
            }
            rows[level - 1].push(column);
            INDEX++;
        });
        return {
            // rowsLen: rows.length,
            // cell,
            // merges,
            // cellLen: INDEX,
            len: {
                r: rows.length,
                c: INDEX
            },
            mergeInfo: merges,
            cellInfo: cellInfo
        };
    } catch (error) {
        console.error(error);
    }
}
//从网络上读取某个excel文件，url必须同域，否则报错
export function httpRequest(url, { responseType = 'arraybuffer', method = 'get' } = { responseType: 'arraybuffer', method: 'get' }) {
    return new Promise((resolve, reject) => {
        let xhr = new XMLHttpRequest();
        xhr.open(method, url, true);
        xhr.responseType = responseType;
        xhr.onload = function () {
            if (xhr.status === 200) {
                resolve(xhr.response);
            }
        };
        xhr.onerror = function (error) {
            return reject(error);
        };
        xhr.send();
    });
}
