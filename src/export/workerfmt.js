/* eslint-disable */
import { isDOM, isObject, isFunction, isArray } from './utils';
import { tableToJson } from './tabletojson';
export function forWokerData(options, useWorker = false) {
    if (!options) {
        return;
    }
    let { table, tables, sheets } = options;
    fcnToString(options, useWorker);
    if (table) {
        tableFcn(table, options, useWorker);
    } else if (tables) {
        tablesFcn(tables, options, useWorker)
    } else if (sheets) {
        sheetsFcn(sheets, options, useWorker);
    }
}
function fcnToString(o, useWorker) {
    if (!useWorker) {
        return;
    }
    for (let item in o) {
        if (isFunction(o[item])) {
            o[item] = o[item].toString();
        }
    }
}
function headerPropFcnToString(header, useWorker) {
    if (!useWorker && !header) {
        return;
    }
    const traverFcn = (item) => {
        let queue = [item];
        while (queue.length) {
            let _header = queue.pop();
            fcnToString(_header, useWorker);
            let { children } = _header;
            if (children && isArray(children)) {
                for (let i = 0, len = children.length; i < len; i++) {
                    queue.push(children[i]);
                }
            }
        }

    };
    if (isArray(header)) {
        for (let i = 0, len = header.length; i < len; i++) {
            traverFcn(header[i]);
        }
    }
}
function tableFcn(table, o, useWorker, key = 'table') {
    if (isDOM(table)) {
        o[key] = { el: tableToJson(table) };
        o[key].el.__isEl = true;
    } else if (isObject(table)) {
        let { el, rowStyle, header } = table;
        headerPropFcnToString(header, useWorker);
        fcnToString(table, useWorker);
        if (isDOM(el)) {
            o[key].el = tableToJson(el, {}, { rowStyle });
            o[key].el.__isEl = true;
        }
    }
}
function tablesFcn(tables, o, useWorker) {
    tablesCallbackFcn(tables, (tbEl, r, c) => {
        tableFcn(tbEl, o.tables[r], useWorker, c);
    });
}
function tablesCallbackFcn(tables, callback) {
    for (let r = 0, rLen = tables.length; r < rLen; r++) {
        let cTables = tables[r];
        for (let c = 0, cLen = cTables.length; c < cLen; c++) {
            let table = cTables[c];
            if (callback) {
                callback(table, r, c);
            }
        }
    }
}
function sheetsFcn(sheets, o, useWorker) {
    for (let i = 0, len = sheets.length; i < len; i++) {
        let { table, tables } = sheets[i];
        fcnToString(sheets[i], useWorker);
        if (table) {
            tableFcn(table, o.sheets[i], useWorker)
        } else if (tables) {
            tablesFcn(tables, o.sheets[i], useWorker);
        }
    }
}