import Excel from 'exceljs';
import { saveAs } from 'file-saver';
import { convertToRows, getAllColumns, typeOf } from './utils';
import { fmtFn } from './fmtfn';
let TMP_MERGECELLINFO = {};//保存单元格合并信息，用于覆盖后面单元格值
export default function exportExcel(opt = {}) {
    let { fileName = '下载', sheets = [], suffixName = "xlsx" } = opt;
    // let workbookConf = {};
    let workbook = createWorkBook();
    initWorkBookViews(workbook);
    createWorkSheets(workbook, sheets);
    toSaveAsFn(workbook, fileName + '.' + suffixName);
}
export function exportExcelUseWorker(opt = {}) {
    return new Promise((resolve, reject) => {
        try {
            let worker = new Worker('./export/export.worker.js');
            fmtFn(opt);
            worker.postMessage({ opt });
            worker.onmessage = function (evt) {
                let { buffer, fileName, success } = evt.data;
                if (success===1) {
                    toSaveAs(buffer, fileName);
                    resolve({ success });
                } else {
                    reject({ success });
                }
            };
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error(error);
            reject(error);
        }
    });
}

function canExport(showTip = true, msg = "该浏览器不支持前端导出功能，请升级浏览器！") {
    let canExp = true;
    if (typeOf(Worker) === 'undefined') {
        canExp = false;
        if (showTip) {
            alert(msg);
        }
    }
    return canExp;
}
export { canExport };
function toSaveAsFn(workbook, fileName) {
    let clearFn = () => {
        TMP_MERGECELLINFO = null;
        workbook = null;
    };
    try {
        getWriteBuffer(workbook).then(buffer => {
            toSaveAs(buffer, fileName);
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
    }
}
function toSaveAs(buffer,fileName) {
    saveAs(new Blob([buffer]), fileName);
}
function getWriteBuffer(workbook) {
    return new Promise((resolve, reject) => {
        workbook.xlsx.writeBuffer()
            .then(buffer => {
                // saveAs(new Blob([buffer]), fileName);
                resolve(buffer);
            }).catch(err => {
                // eslint-disable-next-line no-console
                console.error(err);
                reject(err);
            });
    });
}
function getUrlBase64(img, extension, quality = 1) {
    let canvas = document.createElement("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    let ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, img.width, img.height);
    let ext = extension || img.src.substring(img.src.lastIndexOf(".") + 1).toLowerCase();
    // toDataURL方法，可以是image/jpeg或image/webp,默认image/png
    let dataURL = canvas.toDataURL("image/" + ext, quality);
    return dataURL;
}
async function loadImage(src) {
    return new Promise((resolve, reject) => {
        let img = new Image();
        img.src = src;
        img.onload = function () {
            resolve(img);
        };
        img.onerror = function (error) {
            reject(error);
        };
    });
}
async function getBase64Image(imgSrc, extension = '', quality = 1) {
    let img = await loadImage(imgSrc);
    return getUrlBase64(img, extension, quality);
}
export { getBase64Image };
function createWorkBook() {
    return new Excel.Workbook();
}
function createWorkSheets(workbook, sheets) {
    sheets.forEach((sheet, index) => {
        try {
            TMP_MERGECELLINFO = {};
            let { tables, sheetName = `Sheet${index + 1}`, props, backgroundImage, wsImages } = sheet;
            let worksheet = addWorksheet(workbook, sheetName, props);
            setSheetBackgroundImage(workbook, worksheet, backgroundImage);
            setSheetRangeImages(workbook, worksheet, wsImages);
            createTablesOfSheet(worksheet, tables);
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error(error);
        }
    });
}
async function setSheetBackgroundImage(workbook, worksheet, backgroundImage) {
    try {
        if (backgroundImage && typeOf(backgroundImage) === 'object') {
            let { filename, base64, extension = 'jpeg' } = backgroundImage;
            if (filename) {
                // eslint-disable-next-line no-console
                console.warn('backgroundImage.filename not support！');
            } else if (base64) {
                let imgId = workbook.addImage({
                    base64,
                    extension
                });
                worksheet.addBackgroundImage(imgId);
            }
        }
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function setSheetRangeImages(workbook, worksheet, wsImages) {
    try {
        if (wsImages && typeOf(wsImages) === 'array') {
            wsImages.forEach((image, index) => {
                let { filename, base64, extension = 'jpeg', range } = image;
                if (filename) {
                    // eslint-disable-next-line no-console
                    console.warn(`wsImages[${index}].filename not support！`);
                } else if (base64) {
                    let imgId = workbook.addImage({
                        base64,
                        extension
                    });
                    addSheetImage(worksheet, imgId, range);
                }
            });
        }
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function addSheetImage(worksheet, imgId, opt) {
    try {
        worksheet.addImage(imgId, opt);
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function createTablesOfSheet(worksheet, tables) {
    let tableConf = {
        startC: 0,
        startR: 0,
        maxRowsCount: 0
    };
    tables.forEach(columnTables => {
        tableConf.startC = 0;
        tableConf.maxRowsCount = 0;
        columnTables.forEach(table => {
            createAndAppendTable(worksheet, table, tableConf);
        });
        tableConf.startR += tableConf.maxRowsCount;
    });
}
function getMax(number1, number2) {
    return Math.max(number1, number2);
}
function createAndAppendTable(worksheet, table, tableConf) {
    let { columns = [], data = [], mergeCells, space, origin, rowStyle } = table;
    let { cell, merges, cellLen, rowsLen } = convertToRows(columns);
    let { startC, startR, maxRowsCount } = tableConf;
    let { left = 0, top = 0, bottom = 0, right = 0 } = space || {};
    let hasOrigin = false;
    if (origin && typeOf(origin) === 'object') {
        hasOrigin = true;
        let { r = 0, c = 0 } = origin;
        startC = c > 0 ? c - 1 : 0;
        startR = r > 0 ? r - 1 : 0;
    }

    startC += left;
    startR += top;
    // 头部标题
    cell.forEach(item => {
        let { col, title } = item;
        let _startC = startC + col.c + 1;
        let _startR = startR + col.r + 1;
        worksheet.getCell(_startR, _startC).value = title;
    });
    setHeaderMerge(worksheet, merges, startR, startC);
    // 内容数据
    fillData(worksheet, columns, data, startR + rowsLen, startC, mergeCells, rowStyle);
    let nextHorTableStart = rowsLen + data.length + bottom + top;
    let nextColumnTableStart = cellLen + left + right;
    if (hasOrigin) {
        nextHorTableStart = 0;
        nextColumnTableStart = 0;
    }
    tableConf.startC += nextColumnTableStart;
    tableConf.maxRowsCount = getMax(maxRowsCount, nextHorTableStart);
}
function setBodyMerge(worksheet, mergeCells, startR, startC) {
    let { rowspan = 1, colspan = 1 } = mergeCells;
    let mergeInfo = [];
    if (rowspan > 1 && colspan > 1) {
        mergeInfo = [startR, startC, startR + rowspan - 1, startC + colspan - 1];
    } else if (rowspan > 1) {
        mergeInfo = [startR, startC, startR + rowspan - 1, startC];
    } else if (colspan > 1) {
        mergeInfo = [startR, startC, startR, startC + colspan - 1];
    }
    try {
        worksheet.mergeCells(mergeInfo);
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function fillData(worksheet, columns, data = [], startR, startC, mergeCells, rowStyle) {
    let allColumns = getAllColumns(columns);
    data.forEach((item, rowIndex) => {
        allColumns.forEach((colItem, index) => {
            let { key, cellStyle } = colItem;
            let _startR = rowIndex + 1 + startR;
            let _startC = index + 1 + startC;
            let cell = worksheet.getCell(_startR, _startC);
            let value = getValue(item, colItem, rowIndex) || '';
            cell.value = value;
            if (mergeCells && typeOf(mergeCells) === 'function') {
                let mergeCell = mergeCells({ row: item, rowIndex, key, keyIndex: index });
                if (mergeCell) {
                    TMP_MERGECELLINFO[cell.address] = value;
                    setBodyMerge(worksheet, mergeCell, _startR, _startC);
                }
            }
            if (rowStyle && typeOf(rowStyle) === 'function') {
                setBodyRowStyle(cell, rowStyle, colItem, item, rowIndex);
            }
            if (cellStyle && typeOf(cellStyle) === 'function') {
                // {font,numFmt,alignment,border,fill}
                setBodyCellStyle(cell, cellStyle, colItem, item, rowIndex);
            }
        });
    });
    // 重新覆盖单元格值
    for (let item in TMP_MERGECELLINFO) {
        worksheet.getCell(item).value = TMP_MERGECELLINFO[item];
    }
}
function setBodyRowStyle(cell, rowStyle, col, row, rowIndex) {
    let { key, title, params } = col;
    let style = rowStyle({ row, rowIndex, column: { key, title, params } });
    setBodyStyle(cell, style);
}
function setBodyCellStyle(cell, cellStyle, col, row, rowIndex) {
    let { key, title, params } = col;
    let style = cellStyle({ row, rowIndex, column: { key, title, params } }) || {};
    setBodyStyle(cell, style);
}
function setBodyStyle(cell, style) {
    try {
        let { font, numFmt, alignment, border, fill } = style || {};
        if (font) {
            cell.font = font;
        }
        if (numFmt) {
            cell.numFmt = numFmt;
        }
        if (alignment) {
            cell.alignment = alignment;
        }
        if (border) {
            cell.border = border;
        }
        if (fill) {
            cell.fill = fill;
        }
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function getValue(row, col, rowIndex) {
    try {
        let value = '';
        let { key, fmt, title, params, type } = col;
        value = fmt && typeOf(fmt) === 'function' && fmt({ row, rowIndex, column: { key, title, params } }) || key && row[key] || type && type === 'index' && rowIndex + 1 || '';
        return value;
    } catch (error) {
        // eslint-disable-next-line no-console
        console.error(error);
    }
}
function setHeaderMerge(worksheet, merges = [], startR, startC) {
    merges.forEach(item => {
        try {
            let { s, e } = item;
            worksheet.mergeCells([s.r + 1 + startR, s.c + 1 + startC, e.r + 1 + startR, e.c + 1 + startC]);
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error(error);
        }
    });
}
function addWorksheet(workbook, sheetName, props = {}) {
    return workbook.addWorksheet(sheetName, props);
}
function initWorkBookViews(workbook) {
    workbook.views = [
        {
            x: 0,
            y: 0,
            width: 2000,
            height: 2000,
            firstSheet: 0,
            activeTab: 0,
            visibility: 'visible'
        }
    ];
}