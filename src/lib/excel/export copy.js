/* eslint-disable no-console */
import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';
import { isArray, isDOM, isObject, isFunction, convertToRows, getKeys, encodeAddress2, httpRequest } from './utils';
import { getTableJson } from './table-dom';
import { createSheets } from './sheets';
let workbook = null;
export default function exportExcel(options, configs = {}) {
    try {
        if (!options) {
            console.error('options 不能为空！');
            return;
        }
        initWorkbook();
        initViews();
        createSheets(options.sheets, workbook);
        // let sheetname = 'Sheet1';
        let filename = '下载';
        let suffixName = '.xlsx';
        // if (isDOM(options)) {
        //     createSheetOfTableDom(options);
        // } else if (isObject(options)) {
        //     createSheetOfObj(options);
        // }
        toSaveAsFile(filename + suffixName);
    } catch (error) {
        console.error(error);
    } finally {
        clearFn();
    }
}
export async function readWorkbookFromRemoteFile(url, wk) {
    try {
        let response = await httpRequest(url);
        wk = wk ? wk : new Workbook;
        await wk.xlsx.load(response);
        console.log(wk)
        return wk;
    } catch (error) {
        console.error(error);
    }
}
function createSheetOfObj(options) {
    let { table, tables, sheets, sheetname } = options;
    if (table) {
        if (isDOM(table)) {
            createSheetOfTableDom(table);
        } else if (isObject(table)) {
            createSheetOfTableObj(table);
        }
    } else if (tables) {
        createSheetOfTables(tables);
    } else if (sheets) {
        createSheetsOfSheets();
    }
}
function createSheetOfTableObj(table) {
    // console.log(table)
    let { keys, header, data, sheetname = 'Sheet1', mergeCells, sheetProps = {} } = table;
    let headerJson = convertToRows(header);
    keys = keys ? keys : getKeys(header);
    console.log(keys)
    console.log(headerJson)
    let worksheet = workbook.addWorksheet(sheetname, sheetProps);
    sheetAppendTableJson(worksheet, headerJson);
    sheetAppendData(worksheet, data, keys, mergeCells, { startRow: headerJson.len.r });
}
function createSheetOfTables(tables) {

}
function createSheetsOfSheets(sheets) {
    for (let i = 0, sheetLen = sheets.length; i < sheetLen; i++) {
        let { sheetname, tables } = sheets[i];
        trvTables(tables, sheetname);
    }
}
function trvTables(tables, sheetname) {
    for (let i = 0, len = tables.length; i < len; i++) {
        let rTables = tables[i];
        for (let j = 0, rLen = rTables.length; j < rLen; j++) {
            let table = rTables[j];
            if (isDOM(table)) {

            } else if (isObject(table)) {

            }
        }
    }
}
function createSheetOfTableDom(table, { sheetname, sheetProps } = { sheetname: 'Sheet1', sheetProps: {} }) {
    try {
        let json = getTableJson(table);
        // console.log(json)
        let worksheet = workbook.addWorksheet(sheetname, sheetProps);
        sheetAppendTableJson(worksheet, json);
    } catch (error) {
        console.error(error);
    }
}
function mergeCellsFromFcn(worksheet, sRow, sCol, mergeInfo) {
    try {
        if (!mergeInfo) {
            return;
        }
        let { rowspan = 1, colspan = 1 } = mergeInfo;
        const valueFcn = ({ sr, sc, er, ec }) => {
            worksheet.mergeCells(encodeAddress2(sr, sc), encodeAddress2(er, ec));
        };
        if (rowspan > 1 && colspan === 1) {
            valueFcn({
                sr: sRow,
                sc: sCol,
                er: sRow + rowspan - 1,
                ec: sCol
            });
        } else if (rowspan === 1 && colspan > 1) {
            valueFcn({
                sr: sRow,
                sc: sCol,
                er: sRow,
                ec: sCol + colspan - 1
            });
        } else if (rowspan > 1 && colspan > 1) {
            valueFcn({
                sr: sRow,
                sc: sCol,
                er: sRow + rowspan - 1,
                ec: sCol + colspan - 1
            });
        }
    } catch (error) {
        console.error(error);
    }
}
function mergeCellsFromList(worksheet, mergeInfo) {
    // TODO
}
function sheetAppendData(worksheet, data, keys, mergeCells, { startRow = 0, startCol = 0 } = { startRow: 0, startCol: 0 }) {
    const fcn = (i, row) => {
        for (let k = 0, kLen = keys.length; k < kLen; k++) {
            worksheet.getCell(encodeAddress2(i + startRow, k + startCol)).value = row[keys[k]];
            if (mergeCells && isFunction(mergeCells)) {
                mergeCellsFromFcn(worksheet, i + startRow, k + startCol, mergeCells({ row, rowIndex: i, key: keys[k], keyIndex: k }));
            } else if (mergeCells && isArray(mergeCells)) {
                mergeCellsFromList(worksheet, mergeCells);
            }
        }
    };
    for (let i = 0, dataLen = data.length; i < dataLen; i++) {
        fcn(i, data[i]);
    }
}
function sheetAppendTableJson(worksheet, json) {
    try {
        let { mergeInfo, cellInfo } = json;
        for (let cell in cellInfo) {
            worksheet.getCell(cell).value = cellInfo[cell];
        }
        tableMerge(worksheet, mergeInfo);
    } catch (error) {
        console.error(error);
    }
}
function tableMerge(worksheet, mergeInfo) {
    try {
        for (let i = 0, len = mergeInfo.length; i < len; i++) {
            let { s, e } = mergeInfo[i];
            worksheet.mergeCells(s, e);
        }
    } catch (error) {
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
        activeTab: 0,
        visibility: 'visible'
    };
}
function clearFn() {
    workbook = null;
}
function initWorkbook() {
    workbook = new Workbook();
}
function canExport(showTip = true, msg = "该浏览器不支持前端导出功能，请升级浏览器！") {
    let canExp = true;
    if (typeof Worker === 'undefined') {
        canExp = false;
        if (showTip) {
            alert(msg);
        }
    }
    return canExp;
}
export { canExport };
function toSaveAsFile(fileName) {
    try {
        getWriteBuffer().then(buffer => {
            toSaveAs(buffer, fileName);
        }).catch(e => {
            // eslint-disable-next-line no-console
            console.error(e);
        });
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function toSaveAs(buffer, fileName) {
    try {
        saveAs(new Blob([buffer]), fileName);
    } catch (error) {
        console.error(error);
    }
}
function getWriteBuffer() {
    return new Promise((resolve, reject) => {
        try {
            workbook.xlsx.writeBuffer()
                .then(buffer => {
                    resolve(buffer);
                }).catch(err => {
                    console.error(err);
                    return reject(err);
                });
        } catch (error) {
            console.error(error);
        }
    });
}