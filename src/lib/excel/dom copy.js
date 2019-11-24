import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';
let workbook = null;
export default function exportExcel(sheetOpt, props = {}) {
    try {
        if (!sheetOpt) {
            return;
        }
        workbook = new Workbook();
        initViews();
        let { filename = "下载", sheetname = 'Sheet1', suffixName = '.xlsx' } = props;
        if (isDOM(sheetOpt)) {
            createTableOfDOM(sheetOpt, sheetname);
            toSaveAsFile(filename + suffixName);
            return;
        }
        let { sheets = [] } = sheetOpt;
        filename = sheetOpt.filename ? sheetOpt.filename : filename;
        suffixName = sheetOpt.suffixName ? sheetOpt.suffixName : suffixName;
        createSheets(sheets);
        toSaveAsFile(filename + suffixName);
        // let datas = tableToJson();
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function initViews() {
    workbook.views = {
        x: 0,
        y: 0,
        width: 2000,
        height: 2000,
        firstSheet: 0,
        activeTab: 1,
        visibility: 'visible'
    };
}
function clearFn() {
    workbook = null;
}
function toSaveAsFile(filename) {
    try {
        getWriteBuffer().then(buffer => {
            toSaveAs(buffer, filename);
            clearFn();
        }).catch(e => {
            // eslint-disable-next-line no-console
            console.error(e);
            clearFn();
        });
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
        clearFn();
    } finally {
        clearFn();
    }
}
function toSaveAs(buffer, filename) {
    saveAs(new Blob([buffer]), filename);
}
function getWriteBuffer() {
    return new Promise((resolve, reject) => {
        try {
            workbook.xlsx.writeBuffer()
                .then(buffer => {
                    // saveAs(new Blob([buffer]), fileName);
                    resolve(buffer);
                }).catch(err => {
                    // eslint-disable-next-line no-console
                    console.error(err);
                    reject(err);
                });
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error(error);
            reject(error);
        }
    });
}
function isDOM(str) {
    return str instanceof HTMLElement;
}
function isObject(str) {
    return str instanceof Object;
}
function isArray(str) {
    return str instanceof Array;
}
function createSheets(sheets) {
    try {
        for (let i = 0, len = sheets.length; i < len; i++) {
            let { table, tables } = sheets[i];
            if (table) {
                createTable(table);
            } else if (tables && isArray(tables)) {
                createTables(tables);
            }
            // console.log(isArray(tables))
            // console.log(table instanceof HTMLElement)
            // console.log(tables)
        }
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function createTable(table) {
    if (!table) {
        // eslint-disable-next-line no-console
        console.error('table不能为空！');
        return;
    }
    try {
        if (isDOM(table)) {
            createTableOfDOM(table,'Sheet1');
        } else if (isObject(table)) {
            let { el, keys, columns, data = [], sheetname = 'Sheet1', sheetProps = {} } = table;
            if (el && isDOM(el)) {
                createTableOfDOM(el, sheetname, sheetProps);
            } else if (keys && isArray(keys)) {
                createTableOfKeys(keys, data, sheetname);
            } else if (columns && isArray(columns)) {
                createTableOfColumns(columns, data, sheetname);
            }
        }
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function createTableOfDOM(table, sheetname = 'Sheet1', sheetProps = {}) {
    try {
        let datas = tableToJson(table);
        let worksheet = workbook.addWorksheet(sheetname, sheetProps);
        worksheet.addRows(datas);
          console.log(datas)
        datas = null;
        worksheet = null;
        // let sheet = workbook.add
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function createTableOfKeys() { }
function createTableOfColumns() { }
function createTables(tables) { }
export function tableToJson(el, isarray = true) {
    let datas = [];
    try {
        if (!el) {
            return;
        }
        let trs = el.querySelectorAll('tr');
        for (let i = 0, trsLen = trs.length; i < trsLen; i++) {
            rendCell(trs[i], datas, isarray);
        }
        return datas;
        // console.log(datas)
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    } finally {
        datas = null;
    }
}
function rendCell(tr, datas, isarray) {
    let keys = isarray ? [] : {};
    try {
        let cells = tr.querySelectorAll('td').length && tr.querySelectorAll('td') || tr.querySelectorAll('th');
        // console.log(cells)
        for (let i = 0, cellsLen = cells.length; i < cellsLen; i++) {
            let text = cells[i].innerText;
            // console.log(text)
            if (isarray) {
                keys.push(text);
            } else {
                let key = getFromCharCode(i);
                keys[key] = text;
            }
        }
        datas.push(keys);
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    } finally {
        keys = null;
    }
}
function getFromCharCode(i) {
    return String.fromCharCode(65 + i);
}