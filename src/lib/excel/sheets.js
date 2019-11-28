/* eslint-disable no-console */
import { isDOM, isObject, convertToRows, getKeys, encodeAddress2, isArray, isFunction, isString, decodeAddress } from "./utils";
import { getTableJson } from "./table-dom";

export function createSheets(sheets, workbook) {
    // console.log(sheets)
    for (let s = 0, sLen = sheets.length; s < sLen; s++) {
        let { tables, sheetname, sheetProps = {} } = sheets[s];
        let worksheet = workbook.addWorksheet(sheetname, sheetProps);
        appendTablesToSheet(tables, worksheet);
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
            // console.log(cTables[c]);
            let table = cTables[c];
            console.log(_startCol, 'gg')
            if (isDOM(table)) {
                let json = getTableJson(table, {
                    startRow: startRow,
                    startCol: _startCol
                });
                // opt.startRow += json.len.r + 1;
                cellAppendSheet(json, worksheet);
                _startCol += json.len.c;
                console.log(json)
                _startRow = getMax(_startRow, json.len.r);
                // console.log(json)
            } else if (isObject(table)) {
                let { header, keys, data = [], rowStyle, mergeCells, space = {}, origin } = table;
                console.log(table)
                console.log(_startCol)
                let { left = 0, right = 0, top = 0, bottom = 0 } = origin ? {} : space;
                let { col, row } = getTalbeOrigin(origin);
                console.log(col, row)
                let cTableStartRow = 0;
                if (!origin) {
                    _startCol += left;
                    cTableStartRow = startRow + top;
                }
                let headerData = convertToRows(header, {
                    startRow: !origin ? cTableStartRow : row,
                    startCol: !origin ? _startCol : col
                });
                // opt.startRow += headerData.len.r + 1;
                keys = keys ? keys : getKeys(header);
                cellAppendSheet(headerData, worksheet);
                keysDataAppendSheet(keys, data, rowStyle, mergeCells, worksheet, {
                    startRow: !origin ? cTableStartRow + headerData.len.r : row + headerData.len.r,
                    startCol: !origin ? _startCol : col
                });
                if (origin) {
                    continue;
                }
                _startCol += headerData.len.c + right;
                _startRow = getMax(_startRow, headerData.len.r + data.length + top + bottom);
                console.log(_startRow)
                console.log(headerData)
            }
        }
        startRow = _startRow;
    }
    console.log(startRow)
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
            } else if (mergeCells && isArray(mergeCells)) {
                mergeCellsFromList(worksheet, mergeCells);
            }

            if (rowStyle && isFunction(rowStyle)) {
                rowStyleFromFcn(cell, rowStyle({ row, rowIndex: i, key: keys[k], keyIndex: k }));
            }
            if (isObject(column)) {
                let { cellStyle } = column;
                if (cellStyle && isFunction(cellStyle)) {
                    cellStyleFromFcn(cell,cellStyle({ row, rowIndex: i }));
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
function cellStyleFromFcn(cell,styleInfo) {
    styleFromFcn(cell, styleInfo);
}
function rowStyleFromFcn(cell, styleInfo) {
    styleFromFcn(cell, styleInfo);
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
function mergeCellsFromList(worksheet, mergeInfo) {
    // TODO
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