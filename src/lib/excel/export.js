/* eslint-disable no-console */
import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';
import { isObject, httpRequest } from './utils';
export { getBase64Image } from './image_util';
// import { getTableJson } from './table-dom';
import { createSheets, createSheetTable, createSheetTables, worksheetAddBackgroundImage, worksheetAddImage } from './sheets';
// let workbook = null;
export function exportExcelOfWorker() {
    
}
export default function exportExcel(options, { workbook = initWorkbook() } = { workbook: initWorkbook() }) {
    try {
        if (!options) {
            console.error('options 不能为空！');
            return;
        }
        initViews(workbook);
        let _filename = '下载';
        let _suffixName = '.xlsx';
        if (isObject(options)) {
            let { table, sheetname='sheet1', sheetProps = {}, tables, sheets, filename = '下载', suffixName = '.xlsx', images, backgroundImage } = options;
            _filename = filename;
            _suffixName = suffixName;
            let worksheet = null;
            if (table || tables) {
                worksheet = workbook.addWorksheet(sheetname, sheetProps);
                if (images) {
                    worksheetAddImage(workbook, worksheet, images);
                }
                if (backgroundImage) {
                    worksheetAddBackgroundImage(workbook, worksheet, backgroundImage);
                }
            }
            if (table) {
                createSheetTable(table, worksheet);
            } else if (tables) {
                createSheetTables(tables, worksheet);
            } else if (sheets) {
                createSheets(sheets, workbook);
            }
        }
        toSaveAsFile(workbook, _filename + _suffixName);
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
export function canExport(showTip = true, msg = "该浏览器不支持前端导出功能，请升级浏览器！") {
    let canExp = true;
    if (typeof Worker === 'undefined') {
        canExp = false;
        if (showTip) {
            alert(msg);
        }
    }
    return canExp;
}
function toSaveAsFile(workbook, fileName) {
    try {
        getWriteBuffer(workbook).then(buffer => {
            toSaveAs(buffer, fileName);
        }).catch(e => {
            console.error(e);
        });
    } catch (error) {
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