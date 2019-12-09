/* eslint-disable no-console */
import { isDOM, isObject, convertToRows, getKeys, encodeAddress2, isArray, isFunction, isString, decodeAddress } from "./utils";
import { getTableJson } from "./table-dom";

export function createSheets(sheets, workbook) {
    // console.log(sheets)
    for (let s = 0, sLen = sheets.length; s < sLen; s++) {
        let { tables, sheetname, images, backgroundImage, sheetProps = {} } = sheets[s];
        let worksheet = workbook.addWorksheet(sheetname, sheetProps);
        // appendTablesToSheet(tables, worksheet);
        createSheetTables(tables, worksheet);
        worksheetAddImage(workbook, worksheet, images);
        worksheetAddBackgroundImage(workbook, worksheet, backgroundImage);
    }
}
export function createSheetTables(tables, worksheet) {
    // let worksheet = workbook.addWorksheet(sheetname, sheetProps);
    appendTablesToSheet(tables, worksheet);
}
export function createSheetTable(table, worksheet, config = {}) {
    let { sStartCol: _startCol = 0, startRow = 0, sStartRow: _startRow = 0 } = config;
    let domAppendSheet = (el, sRow = 0, sCol = 0) => {
        let json = getTableJson(el, {
            startRow: sRow,
            startCol: sCol
        });
        // opt.startRow += json.len.r + 1;
        cellAppendSheet(json, worksheet);
        // _startCol += json.len.c;
        // _startRow = getMax(sRow, json.len.r);
        let { len } = json;
        return {
            startCol: len.c,
            startRow: len.r + sRow,
            len: len
        };
    };
    // let dLenC = 0;
    if (isDOM(table)) {
        // console.log(json)
        let l = domAppendSheet(table);
        _startCol = l.startCol;
        _startRow = getMax(_startRow, l.startRow);
    } else if (isObject(table)) {
        let { el, header, keys, data = [], rowStyle, rowStyleList, mergeCells, mergeCellsList, space = {}, origin } = table;
        let { left = 0, right = 0, top = 0, bottom = 0 } = origin ? {} : space;
        let { col, row } = getTalbeOrigin(origin);
        let cTableStartRow = 0;
        if (!origin) {
            _startCol += left;
            cTableStartRow = startRow + top;
        }
        let hLenC = 0;
        let hLenR = 0;
        const stylefcn = (dLenC, sRow, sCol) => {
            // console.log(!rowStyle && rowStyleList && isArray(rowStyleList), dLenC)
            let r = !origin ? sRow : row;
            let c = !origin ? sCol : col;
            console.log(dLenC, r, c);
            if (!rowStyle && rowStyleList && isArray(rowStyleList)) {
                rowStyleFromList(worksheet, rowStyleList, dLenC, r, c);
            }
            // if (!mergeCells && mergeCellsList && isArray(mergeCellsList)) {
            //     mergeCellsFromList(worksheet, mergeCellsList, r, c,dom);
            // }
        };
        const mergeFcn = (sRow, sCol) => {
            let r = !origin ? sRow : row;
            let c = !origin ? sCol : col;
            if (!mergeCells && mergeCellsList && isArray(mergeCellsList)) {
                mergeCellsFromList(worksheet, mergeCellsList, r, c);
            }
        };
        if (el && isDOM(el)) {
            let l = domAppendSheet(el, cTableStartRow, _startCol);
            stylefcn(l.len.c, cTableStartRow + l.len.header, _startCol);
            mergeFcn(cTableStartRow + l.len.header, _startCol);
            _startCol = l.startCol;
            _startRow = getMax(_startRow, l.startRow);
        } else {
            let headerData = convertToRows(header, {
                startRow: !origin ? cTableStartRow : row,
                startCol: !origin ? _startCol : col
            });
            hLenC = headerData.len.c;
            hLenR = headerData.len.r;
            keys = keys ? keys : getKeys(header);
            // fcn(keys.length, cTableStartRow + hLenR, _startCol);
            stylefcn(keys.length, cTableStartRow + hLenR, _startCol);
            mergeFcn(cTableStartRow + hLenR, _startCol);
            cellAppendSheet(headerData, worksheet);
            keysDataAppendSheet(keys, data, rowStyle, mergeCells, worksheet, {
                startRow: !origin ? cTableStartRow + hLenR : row + hLenR,
                startCol: !origin ? _startCol : col
            });
        }

        if (!origin) {
            _startCol += hLenC + right;
            _startRow = getMax(_startRow, hLenR + data.length + top + bottom);
        }
    }
    return {
        sStartCol: _startCol,
        sStartRow: _startRow
    };
}
function workbookAddBase64Image(img, extension = 'png', workbook) {
    return workbook.addImage({
        base64: img,
        extension: extension,
    });
}
export function worksheetAddImage(workbook, worksheet, images) {
    try {
        if (images && isArray(images)) {
            // let worksheet = workbook.getWorksheet(sheetname);
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
    } catch (error) {
        console.error(error);
    }
}
export function worksheetAddBackgroundImage(workbook, worksheet, backgroundImage) {
    try {
        if (!backgroundImage) {
            return;
        }
        let { base64, extension = 'png' } = backgroundImage;
        // let worksheet = workbook.getWorksheet(sheetname);
        let img = workbookAddBase64Image(base64, extension, workbook);
        worksheet.addBackgroundImage(img);
    } catch (error) {
        console.error(error);
    }
}
function appendTablesToSheet(tables, worksheet) {
    // let opt = {
    //     startRow: 0,
    //     startCol: 0
    // };
    let startRow = 0;
    for (let r = 0, rLen = tables.length; r < rLen; r++) {
        let cTables = tables[r];
        let _startCol = 0;
        let _startRow = startRow;
        for (let c = 0, cLen = cTables.length; c < cLen; c++) {
            let table = cTables[c];
            let { sStartRow, sStartCol } = createSheetTable(table, worksheet, {
                sStartCol: _startCol,
                sStartRow: _startRow,
                startRow: startRow
            });
            _startCol = sStartCol;
            _startRow = sStartRow;
        }
        startRow = _startRow;
    }
}
function cellAppendSheet(data, worksheet) {
    // console.log(data)
    try {
        let { mergeInfo, cellInfo } = data;
        for (let cell in cellInfo) {
            worksheet.getCell(cell).value = cellInfo[cell];
        }
        tableMerge(mergeInfo, worksheet);
    } catch (error) {
        console.error(error);
    }
}
// return {col:,row:}
function getTalbeOrigin(origin) {
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
function keysDataAppendSheet(keys, data, rowStyle, mergeCells, worksheet, { startRow = 0, startCol = 0 } = { startRow: 0, startCol: 0 }) {
    const fcn = (i, row) => {
        for (let k = 0, kLen = keys.length; k < kLen; k++) {
            let cell = worksheet.getCell(encodeAddress2(i + startRow, k + startCol));
            let column = keys[k];
            cell.value = getCellValue(row, column, i + startRow, k + startCol);

            if (mergeCells && isFunction(mergeCells)) {
                mergeCellsFromFcn(worksheet, i + startRow, k + startCol, mergeCells({ row, rowIndex: i, key: keys[k], keyIndex: k }));
            }

            if (rowStyle && isFunction(rowStyle)) {
                rowStyleFromFcn(cell, rowStyle({ row, rowIndex: i, key: keys[k], keyIndex: k }));
            }
            if (isObject(column)) {
                let { cellStyle } = column;
                if (cellStyle && isFunction(cellStyle)) {
                    cellStyleFromFcn(cell, cellStyle({ row, rowIndex: i }));
                }
            }
        }
    };
    for (let i = 0, dataLen = data.length; i < dataLen; i++) {
        fcn(i, data[i]);
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
function cellStyleFromFcn(cell, styleInfo) {
    styleFromFcn(cell, styleInfo);
}
function rowStyleFromFcn(cell, styleInfo) {
    styleFromFcn(cell, styleInfo);
}
function rowStyleFromList(worksheet, styleList, keysLen, startRow, startCol) {
    console.log(styleList)
    for (let i = 0, len = styleList.length; i < len; i++) {
        let { style, index } = styleList[i];
        for (let k = 0; k < keysLen; k++) {
            let cell = worksheet.getCell(encodeAddress2(index + startRow, k + startCol));
            styleFromFcn(cell, style);
        }
    }
}
function styleFromFcn(cell, styleInfo) {
    if (!styleInfo) {
        return;
    }
    for (let styleItem in styleInfo) {
        cell[styleItem] = styleInfo[styleItem];
    }
}
function mergeCellsFromFcn(worksheet, sRow, sCol, mergeInfo) {
    try {
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
    } catch (error) {
        console.error(error);
    }
}
function mergeCellsFromList(worksheet, mergeList, sRow, sCol) {
    for (let i = 0, len = mergeList.length; i < len; i++) {
        let { rowIndex, keyIndex, rowspan, colspan } = mergeList[i];
        let mergeInfo = {
            rowspan,
            colspan
        };
        mergeCellsFromFcn(worksheet, sRow + rowIndex, sCol + keyIndex, mergeInfo);
    }
}
function tableMerge(mergeInfo, worksheet) {
    try {
        for (let i = 0, len = mergeInfo.length; i < len; i++) {
            let { s, e } = mergeInfo[i];
            worksheetMergeCells(worksheet, s, e);
        }
    } catch (error) {
        console.error(error);
    }
}
function worksheetMergeCells(worksheet, s, e) {
    worksheet.mergeCells(s, e);
}

function getMax(a, b) {
    return Math.max(a, b);
}