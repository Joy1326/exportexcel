/* eslint-disable */

import { isEmpty, isFunction, encodeAddress2, n2l } from './utils';
export function tableToJson(table, config = {}, options = {}) {
    try {
        let { startRow = 0, startCol = 0 } = config;
        let headJson = getTableHeaderJson(table, { startRow, startCol });
        let bodyJson = getTableBodyJson(table, { startRow: headJson.len.r + startRow, startCol }, options);
        return {
            len: {
                header: headJson.len.r,
                body: bodyJson.len.r,
                r: headJson.len.r + bodyJson.len.r,
                c: headJson.len.c ? headJson.len.c : bodyJson.len.c
            },
            cellInfo: { ...bodyJson.cellInfo, ...headJson.cellInfo },
            mergeInfo: [...bodyJson.mergeInfo, ...headJson.mergeInfo],
            cellStyleInfo: { ...bodyJson.cellStyleInfo }
        };
    } catch (error) {
        console.error(error);
    }
}
function getTableHeaderJson(table, config, options) {
    try {
        let thead = table.querySelector('thead');
        return getTableCellInfo(thead, config, options);
    } catch (error) {
        console.error(error);
    }
}
function getTableBodyJson(table, config, options) {
    try {
        let tbody = table.querySelector('tbody');
        return getTableCellInfo(tbody, config, options);
    } catch (error) {
        console.error(error);
    }
}
function getTableCellInfo(el, config, options = {}) {
    let cellOpt = {
        cellInfo: {},
        cellStyleInfo: {},
        mergeInfo: [],
        len: {
            r: 0,
            c: 0
        }
    };
    if (!el) {
        return cellOpt;
    }
    try {
        let trs = el.querySelectorAll('tr');
        let { rowStyle,cellStyle } = options;
        let opt = {
            k: 0
        };
        let trLen = trs.length;
        let row = null;
        let tag = rowStyle && isFunction(rowStyle);
        for (let i = 0; i < trLen; i++) {
            console.log();
            let cells = getCellEls(trs[i]);
            if (tag) {
                row = getRowData(cells);
            }
            let rStyle = getCellAttrStyle(trs[i].getAttribute('data-row-style'));
            // console.log(row)
            opt.k = 0;
            for (let k = 0, cellsLen = cells.length; k < cellsLen; k++) {
                let cell = cells[k];
                opt.k = opt.k > k ? opt.k : k;
                let { cellName, cellStyle } = fillCellValue(cell, cellOpt, i, opt, config);
                if (rStyle) {
                    rowStyleFcn(cellOpt.cellStyleInfo, cellName, rStyle);
                }
                if (tag) {
                    let style = rowStyle({ row: row, rowIndex: i, key: n2l(k + 1), keyIndex: k });
                    // console.log(cellName,cell)
                    if (style) {
                        rowStyleFcn(cellOpt.cellStyleInfo, cellName, style);
                    }
                }
                if (cellStyle) {
                    rowStyleFcn(cellOpt.cellStyleInfo, cellName, cellStyle);
                }
                opt.k++;
            }
            if (i === 0) {
                cellOpt.len.r = trLen;
                cellOpt.len.c = opt.k;
            }
        }
        opt = null;
        return cellOpt;
    } catch (error) {
        console.error(error);
    } finally {
        cellOpt = null;
    }
}
function getCellAttrStyle(style) {
    if (!style) {
        return;
    }
    try {
        return JSON.parse(style);
    } catch (error) {
        console.error(error);
    }
}
function rowStyleFcn(cellStyleInfo, cellName, style) {
    cellStyleInfo[cellName] = {};
    Object.keys(style).forEach(item => {
        cellStyleInfo[cellName][item] = style[item];
    });
}
function getRowData(cells) {
    let row = {};
    try {
        for (let k = 0, cellsLen = cells.length; k < cellsLen; k++) {
            row[n2l(k + 1)] = getCellElText(cells[k]);
        }
        return row;
    } catch (error) {
        console.error(error);
    } finally {
        row = null;
    }
}
function fillCellValue(cell, cellOpt, rowIndex, opt, { startRow = 0, startCol = 0 } = { startRow: 0, startCol: 0 }) {
    let { rowspan, colspan } = getCellElSpan(cell);
    let text = getCellElText(cell);
    let { cellInfo, mergeInfo } = cellOpt;
    const mergeName = 'ly-merge';
    let cellName = null;
    const mergeFcn = ({ sr, sc, er, ec }) => {
        mergeInfo.push({
            s: encodeAddress2(sr + startRow, sc + startCol),
            e: encodeAddress2(er + startRow, ec + startCol)
        });
    };
    const valueFcn = (r, c, tag = true) => {
        r = r + startRow;
        c = c + startCol;
        let address1 = encodeAddress2(r, c);
        if (isEmpty(cellInfo[address1])) {
            if (tag) {
                cellInfo[address1] = text;
                cellName = address1;
            } else {
                cellInfo[address1] = mergeName;
            }
        } else {
            opt.k++;
            while (!isEmpty(cellInfo[encodeAddress2(r, opt.k + startCol)])) {
                opt.k++;
            }
            let address2 = encodeAddress2(r, opt.k + startCol);
            cellInfo[address2] = text;
            cellName = address2;
        }
    };
    const cFcn = () => {
        let { k } = opt;
        mergeFcn({
            sr: rowIndex,
            sc: k,
            er: rowIndex,
            ec: k + colspan - 1
        });
        for (let c = opt.k; c < opt.k + colspan; c++) {
            valueFcn(rowIndex, c, c === k);
        }
    };
    const rFcn = () => {
        mergeFcn({
            sr: rowIndex,
            sc: opt.k,
            er: rowIndex + rowspan - 1,
            ec: opt.k
        });
        for (let r = rowIndex; r < rowIndex + rowspan; r++) {
            valueFcn(r, opt.k, r === rowIndex);
        }
    };
    const rcFcn = () => {
        let { k } = opt;
        mergeFcn({
            sr: rowIndex,
            sc: k,
            er: rowIndex + rowspan - 1,
            ec: k + colspan - 1
        });
        for (let r = rowIndex; r < rowIndex + rowspan; r++) {
            for (let c = opt.k; c < opt.k + colspan; c++) {
                valueFcn(r, c, r === rowIndex && c === k);
            }
        }
    };
    if (rowspan === 1 && colspan === 1) {
        valueFcn(rowIndex, opt.k);
    } else if (rowspan > 1 && colspan === 1) {
        rFcn();
    } else if (rowspan === 1 && colspan > 1) {
        cFcn();
    } else if (rowspan > 1 && colspan > 1) {
        rcFcn();
    }
    return {
        cellName,
        cellStyle: getCellAttrStyle(cell.getAttribute('data-cell-style'))
    };
}
function getCellElText(cell) {
    return cell.innerText;
}
function getCellElSpan(cell) {
    return {
        rowspan: getCellElRowspan(cell),
        colspan: getCellElColspan(cell)
    };
}
function getCellElRowspan(cell) {
    return +cell.getAttribute('rowspan') || 1;
}
function getCellElColspan(cell) {
    return +cell.getAttribute('colspan') || 1;
}
function getCellEls(tr) {
    return tr.querySelectorAll('td').length && tr.querySelectorAll('td') || tr.querySelectorAll('th');
}