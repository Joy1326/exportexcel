/* eslint-disable no-console */
import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';
import { isArray, isDOM, isObject, isFunction, convertToRows, getKeys, encodeAddress2, httpRequest } from './utils';
import { getTableJson } from './table-dom';
import { createSheets } from './sheets';
// let workbook = null;
export default function exportExcel(options, { workbook = initWorkbook() } = { workbook: initWorkbook() }) {
    try {
        if (!options) {
            console.error('options 不能为空！');
            return;
        }
        initViews(workbook);
        createSheets(options.sheets, workbook);
        console.log(workbook)
        // let sheetname = 'Sheet1';
        let filename = '下载';
        let suffixName = '.xlsx';
        toSaveAsFile(workbook, filename + suffixName);
    } catch (error) {
        console.error(error);
    } finally {
        workbook = null;
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

function initViews(workbook) {
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
function initWorkbook() {
    return new Workbook();
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
function toSaveAsFile(workbook, fileName) {
    try {
        getWriteBuffer(workbook).then(buffer => {
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
function getWriteBuffer(workbook) {
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