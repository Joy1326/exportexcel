/* eslint-disable no-console */
import { Workbook } from 'exceljs';

import { isObject, isArray, isFunction, decodeAddress, encodeAddress2, isString, convertToRows, getKeys, httpRequest } from './utils';
if(self){
    self.__exportExcel = exportExcel;
}
export default function exportExcel(options) {
    return new Promise(async (resolve, reject) => {
        if (!options) {
            console.error('options不能为空！');
            return reject(new Error('options不能为空！'));
        }
        if (!isObject(options)) {
            console.error('options必须是object！');
            return reject(new Error('options必须是object！'));
        }
        let workbook = initWorkBook();
        let { table, tables, sheets, xlsxFile, sheetname = getNextSheetname(workbook),  sheetProps = {}, images, backgroundImage } = options;
        if (xlsxFile) {
            workbook = await readWorkbookFromRemoteFile(xlsxFile, workbook);
        }
        if (!xlsxFile) {
            initViews(workbook);
        }
        let worksheet = null;
        if (table || tables) {
            worksheet = createWorkSheet(workbook, sheetname, sheetProps);
        }
        if (table) {
            tableFcn({
                worksheet,
                table
            });
        } else if (tables) {
            tablesFcn({
                worksheet,
                tables
            });
        } else if (sheets && isArray(sheets)) {
            sheetsFcn({
                sheets,
            }, workbook);
        }
        worksheetAddImages(images, workbook, worksheet);
        worksheetAddBackgroundImage(backgroundImage, workbook, worksheet);
        if (table || tables || sheets || xlsxFile) {
            resolve(await getWriteBuffer(workbook));
        }
    });
}
function worksheetAddImages(images, workbook, worksheet) {
    if (images && isArray(images)) {
        for (let i = 0, len = images.length; i < len; i++) {
            let { base64, extension = 'png', range } = images[i];
            if (base64 && range) {
                let img = workbookAddBase64Image(base64, extension, workbook);
                worksheet.addImage(img, range);
            } else {
                console.error(`index${i} base64和range必须传值`);
            }
        }
    }
}
function worksheetAddBackgroundImage(backgroundImage, workbook, worksheet) {
    if (!backgroundImage) {
        return;
    }
    let { base64, extension = 'png' } = backgroundImage;
    let img = workbookAddBase64Image(base64, extension, workbook);
    worksheet.addBackgroundImage(img);
}
function workbookAddBase64Image(img, extension = 'png', workbook) {
    return workbook.addImage({
        base64: img,
        extension: extension,
    });
}
function getNextSheetname(workbook) {
    if (workbook && workbook.nextId) {
        return 'sheet' + workbook.nextId;
    }
    return 'sheet1';
}
function tableElAppendSheet(data, worksheet, config = {}) {
    let { startCol = 0, startRow = 0 } = config;
    appendCellDatasToSheet(data, worksheet, { startCol, startRow });
    return {
        r: data.len.r,
        c: data.len.c
    };
}
function appendCellDatasToSheet(data, worksheet, config = {}) {
    let { cellInfo, mergeInfo, cellStyleInfo } = data;
    let { startCol = 0, startRow = 0 } = config;
    Object.keys(cellInfo).forEach(cell => {
        let newCell = getNewCell(cell, { r: startRow, c: startCol });
        let wsCell = worksheet.getCell(newCell);
        wsCell.value = cellInfo[cell];
        if (cellStyleInfo && cellStyleInfo[cell]) {
            styleFcn(wsCell, cellStyleInfo[cell]);
        }
    });
    tableMergeCells(mergeInfo, worksheet, { startCol, startRow });
}
function getNewCell(cell, { c = 0, r = 0 } = { r, c }) {
    let { col, row } = decodeAddress(cell);
    return encodeAddress2(row + r - 1, col + c - 1);
}
function tableMergeCells(mergeInfo, worksheet, { startCol, startRow }) {
    for (let i = 0, len = mergeInfo.length; i < len; i++) {
        let { s, e } = mergeInfo[i];
        let config = {
            r: startRow,
            c: startCol
        };
        s = getNewCell(s, config);
        e = getNewCell(e, config);
        worksheetMergeCells(worksheet, s, e);
    }
}
function worksheetMergeCells(worksheet, s, e) {
    worksheet.mergeCells(s, e);
}
function keysDataAppendSheet(options, worksheet, { startCol, startRow }) {
    let { header, keys, data = [], rowStyle, mergeCells } = options;
    let headerData = convertToRows(header, {
        startCol: startCol,
        startRow: startRow
    });
    keys = keys ? keys : getKeys(header);
    let { r, c } = headerData.len;
    appendCellDatasToSheet(headerData, worksheet);
    dataListAppendSheet(worksheet, {
        data,
        keys,
        mergeCells,
        rowStyle
    }, {
        startCol: startCol,
        startRow: startRow + r
    });
    return {
        r: r + data.length,
        c: c
    };
}
function dataListAppendSheet(worksheet, options, { startCol, startRow }) {
    let { data = [], keys, mergeCells, rowStyle } = options;
    const fcn = (i, row) => {
        for (let k = 0, kLen = keys.length; k < kLen; k++) {
            let cell = worksheet.getCell(encodeAddress2(i + startRow, k + startCol));
            let column = keys[k];
            cell.value = getCellValue(row, column, i + startRow, k + startCol);

            if (mergeCells && isFunction(mergeCells)) {
                mergeCellsFcn(worksheet, i + startRow, k + startCol, mergeCells({ row, rowIndex: i, key: keys[k], keyIndex: k }));
            }

            if (rowStyle && isFunction(rowStyle)) {
                rowStyleFcn(cell, rowStyle({ row, rowIndex: i, key: keys[k], keyIndex: k }));
            }
            if (isObject(column)) {
                let { cellStyle } = column;
                if (cellStyle && isFunction(cellStyle)) {
                    cellStyleFcn(cell, cellStyle({ row, rowIndex: i }));
                }
            }
        }
    };
    for (let i = 0, len = data.length; i < len; i++) {
        fcn(i, data[i]);
    }
}
function mergeCellsFcn(worksheet, sRow, sCol, mergeInfo) {
    if (!mergeInfo) {
        return;
    }
    let { rowspan = 1, colspan = 1 } = mergeInfo;
    const valueFcn = ({ sr, sc, er, ec }) => {
        worksheetMergeCells(worksheet, encodeAddress2(sr, sc), encodeAddress2(er, ec));
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
}
function rowStyleFcn(cell, styleInfo) {
    styleFcn(cell, styleInfo);
}
function cellStyleFcn(cell, styleInfo) {
    styleFcn(cell, styleInfo);
}
function styleFcn(cell, styleInfo) {
    if (!styleInfo) {
        return;
    }
    for (let styleItem in styleInfo) {
        cell[styleItem] = styleInfo[styleItem];
    }
}
function getCellValue(row, column, sRow, sCol) {
    if (isObject(column)) {
        let { key, fmt } = column;
        if (fmt && isFunction(fmt)) {
            return fmt({ row, rowIndex: sRow, key, keyIndex: sCol });
        }
        return row[key] || '';
    } else if (isString(column)) {
        return row[column];
    }
    return '';
}
function tableFcn(options) {
    let { table, worksheet } = options;
    let { col = 0, row = 0 } = options.origin || {};
    let nRow = row;
    let nCol = col;
    let { el, space = {}, origin } = table;
    let { left = 0, top = 0, right = 0, bottom = 0 } = space;
    let _origin = getOrigin(origin);
    let _startCol = col + left;
    let _startRow = row + top;
    if (origin) {
        _startCol = _origin.col + left;
        _startRow = _origin.row + top;
    }
    let L = null;
    if (isObject(el)) {
        L = tableElAppendSheet(el, worksheet, {
            startCol: _startCol,
            startRow: _startRow
        });
    } else {
        L = keysDataAppendSheet(table, worksheet, {
            startCol: _startCol,
            startRow: _startRow
        });
    }
    if (!origin) {
        nRow = L.r + _startRow + bottom;
        nCol = L.c + _startCol + right;
    }
    return {
        sRow: nRow,
        sCol: nCol
    };
}
function tablesFcn(options) {
    let { worksheet, tables } = options;
    let nextStartRow = 0;
    let nextStartCol = 0;
    let _nextStartRow = 0;
    tablesCallbackFcn(tables, {
        before: () => {
            nextStartCol = 0;
            _nextStartRow = nextStartRow;
        },
        enter: (table) => {
            let { sCol, sRow } = tableFcn({
                worksheet,
                table: table,
                origin: {
                    col: nextStartCol,
                    row: nextStartRow
                }
            });
            nextStartCol = sCol;
            _nextStartRow = getMax(_nextStartRow, sRow);
        },
        after: () => {
            nextStartRow = _nextStartRow;
        }
    });
}
function tablesCallbackFcn(tables, options = {}) {
    let { before, enter, after } = options;
    for (let r = 0, rLen = tables.length; r < rLen; r++) {
        let cTables = tables[r];
        if (before) {
            before(r);
        }
        for (let c = 0, cLen = cTables.length; c < cLen; c++) {
            let table = cTables[c];
            if (enter) {
                enter(table, r, c);
            }
        }
        if (after) {
            after(r);
        }
    }
}
function sheetsFcn(options, workbook) {
    let { sheets } = options;
    for (let i = 0, len = sheets.length; i < len; i++) {
        let { sheetname = getNextSheetname(workbook), sheetProps = {}, table, tables, images, backgroundImage } = sheets[i];
        let worksheet = createWorkSheet(workbook, sheetname, sheetProps);
        if (table) {
            tableFcn({
                worksheet,
                table
            });
        } else if (tables) {
            tablesFcn({
                worksheet,
                tables
            });
        }
        worksheetAddImages(images, workbook, worksheet);
        worksheetAddBackgroundImage(backgroundImage, workbook, worksheet);
    }
}
function createWorkSheet(workbook, sheetname, sheetProps = {}) {
    if (workbook.getWorksheet(sheetname)) {
        return workbook.getWorksheet(sheetname);
    }
    return workbook.addWorksheet(sheetname, sheetProps);
}
function initWorkBook() {
    return new Workbook();
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
function getMax(a, b) {
    return Math.max(a, b);
}
// return {col:,row:}
function getOrigin(origin) {
    if (!origin) {
        return {
            col: 0,
            row: 0
        };
    }
    if (isString(origin)) {
        let { col, row } = decodeAddress(origin);
        return {
            col: col - 1,
            row: row - 1
        };
    }
    return origin;
}
export async function readWorkbookFromRemoteFile(url, workbook) {
    workbook = workbook ? workbook : new Workbook;
    try {
        let response = await httpRequest(url);
        await workbook.xlsx.load(response);
        return workbook;
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
