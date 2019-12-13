/* eslint-disable */
import { isDOM, isObject } from './utils';
import { tableToJson } from './tabletojson';
export function forWokerData(options) {
    if (!options) {
        return;
    }
    let { table, tables, sheets } = options;
    if (table) {
        tableFcn(table, options);
    } else if (tables) {
        tablesFcn(tables, options)
    } else if (sheets) {
        sheetsFcn(sheets, options);
    }
}
function tableFcn(table, o) {
    if (isDOM(table)) {
        o.table = { el: tableToJson(table) };
        o.table.el.__isEl = true;
    } else if (isObject(table)) {
        let { el, rowStyle } = table;
        if (isDOM(el)) {
            o.table.el = tableToJson(el, {}, { rowStyle });
            o.table.el.__isEl = true;
        }
    }
}
function tablesFcn(tables, o) {
    tablesCallbackFcn(tables, (tbEl, r, c) => {
        if (isDOM(tbEl)) {
            o.tables[r][c] = { el: tableToJson(tbEl) };
            o.tables[r][c].el.__isEl = true;
        } else if (isObject(tbEl)) {
            let { el, rowStyle } = tbEl;
            if (isDOM(el)) {
                o.tables[r][c].el = tableToJson(el, {}, { rowStyle });
                o.tables[r][c].el.__isEl = true;
            }
        }
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
function sheetsFcn(sheets, o) {
    for (let i = 0, len = sheets.length; i < len; i++) {
        let { table, tables } = sheets[i];
        if (table) {
            if (isDOM(table)) {
                o.sheets[i].table = {
                    el: tableToJson(table),
                };
                o.sheets[i].table.el.__isEl = true;
            } else if (isObject(table)) {
                let { el, rowStyle } = table;
                if (isDOM(el)) {
                    o.sheets[i].table.el = tableToJson(el, {}, { rowStyle });
                    o.sheets[i].table.el.__isEl = true;
                }
            }
        } else if (tables) {
            tablesCallbackFcn(tables, (tbEl, r, c) => {
                if (isDOM(tbEl)) {
                    o.sheets[i].tables[r][c] = { el: tableToJson(tbEl) };
                    o.sheets[i].tables[r][c].el.__isEl = true;
                } else if (isObject(tbEl)) {
                    let { el, rowStyle } = tbEl;
                    if (isDOM(el)) {
                        o.sheets[i].tables[r][c].el = tableToJson(el, {}, { rowStyle });
                        o.sheets[i].tables[r][c].el.__isEl = true;
                    }
                }
            });
        }
    }
}