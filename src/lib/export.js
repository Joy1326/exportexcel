import Excel from 'exceljs';
import { saveAs } from 'file-saver';
// import Range from 'exceljs/dist/es5/doc/range';
import { convertToRows, getAllColumns,typeOf } from './utils';
// window.Range = Range;
let TMP_MERGECELLINFO = {};//保存单元格合并信息，用于覆盖后面单元格值
function exportExcell(opt = {}) {
    let { fileName = '下载', sheets = [], suffixName = "xlsx" } = opt;
    // let workbookConf = {};
    let workbook = createWorkBook();
    initWorkBookViews(workbook);
    createWorkSheets(workbook, sheets);
    toSaveAs(workbook, fileName + '.' + suffixName);
}
function toSaveAs(workbook, fileName) {
    let clearFn = () => {
        TMP_MERGECELLINFO = null;
        workbook = null;
    };
    workbook.xlsx.writeBuffer()
        .then(buffer => {
            saveAs(new Blob([buffer]), fileName);
            clearFn();
        }).catch(err => {
            // eslint-disable-next-line no-console
            console.error(err);
            clearFn();
        });
}
function createWorkBook() {
    return new Excel.Workbook();
}
function createWorkSheets(workbook, sheets) {
    sheets.forEach((sheet, index) => {
        try {
            TMP_MERGECELLINFO = {};
            let { tables, sheetName = `Sheet${index + 1}`,props } = sheet;
            let worksheet = addWorksheet(workbook, sheetName,props);
            createTablesOfSheet(worksheet, tables);
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error(error);
        }
    });
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
    let { columns = [], data = [], mergeCells, space, origin } = table;
    let { cell, merges, cellLen, rowsLen } = convertToRows(columns);
    let { startC, startR, maxRowsCount } = tableConf;
    let { left = 0, top = 0, bottom = 0, right = 0 } = space || {};
    let hasOrigin = false;
    if (origin&&typeOf(origin)==='object') {
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
    fillData(worksheet, columns, data, startR + rowsLen, startC, mergeCells);
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
function fillData(worksheet, columns, data = [], startR, startC, mergeCells) {
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
            if (cellStyle && typeOf(cellStyle) === 'function') {
                // {font,numFmt,alignment,border,fill}
                setBodyStyle(cell, cellStyle, colItem, item, rowIndex);
            }
        });
    });
    // 重新覆盖单元格值
    for (let item in TMP_MERGECELLINFO) {
        worksheet.getCell(item).value = TMP_MERGECELLINFO[item];
    }
}
function setBodyStyle(cell, cellStyle, col, row, rowIndex) {
    try {
        let { key, title, params } = col;
        let { font, numFmt, alignment, border, fill } = cellStyle({ row, rowIndex, column: { key, title, params } }) || {};
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
function addWorksheet(workbook, sheetName,props={}) {
    return workbook.addWorksheet(sheetName,props);
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
export {
    exportExcell as default
};
