/* eslint-disable no-console */
import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';
import { isArray, isDOM, isObject } from './utils';
import { getTableJson } from './table-dom';
let workbook = null;
export default function exportExcel(options, configs = {}) {
    try {
        if (!options) {
            console.error('options 不能为空！');
            return;
        }
        initWorkbook();
        initViews();
        if (isDOM(options)) {
            createSheetOfTableDom(options);
        }
        toSaveAsFile('ss.xlsx');
    } catch (error) {
        console.error(error);
    } finally {
        clearFn();
    }
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
function createSheetOfTableDom(table, { sheetname, sheetProps } = { sheetname: 'Sheet1', sheetProps: {} }) {
    try {
        let json = getTableJson(table);
        console.log(json)
        let worksheet = workbook.addWorksheet(sheetname, sheetProps);
        sheetAppendTableJson(worksheet, json);
    } catch (error) {
        console.error(error);
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
function tableMerge(worksheet,mergeInfo) {
    try {
        for (let i = 0, len = mergeInfo.length; i < len; i++){
            let { s, e } = mergeInfo[i];
            worksheet.mergeCells(s,e);
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