/* eslint-disable no-console */
import { isEmpty,encodeAddress2 } from './utils';
export function getTableJson(table,{startRow=0,startCol=0}={startRow:0,startCol:0}) {
    try {
        let headJson = getTableHeaderJson(table,{startRow,startCol});
        let bodyJson = getTableBodyJson(table, { startRow: headJson.len.r+startRow,startCol });
        return {
            len: {
                r: headJson.len.r + bodyJson.len.r,
                c: headJson.len.c?headJson.len.c:bodyJson.len.c
            },
            cellInfo: { ...bodyJson.cellInfo, ...headJson.cellInfo },
            mergeInfo: [...bodyJson.mergeInfo, ...headJson.mergeInfo]
        };
    } catch (error) {
        console.error(error);
    }
}
function getTableHeaderJson(table, opt) {
    try {
        let thead = table.querySelector('thead');
        return getTableCellInfo(thead, opt);
    } catch (error) {
        console.error(error);
    }
}
function getTableBodyJson(table, opt) {
    try {
        let tbody = table.querySelector('tbody');
        return getTableCellInfo(tbody, opt);
    } catch (error) {
        console.error(error);
    }
}
function getTableCellInfo(el, config) {
    let cellOpt = {
        cellInfo: {},
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
        let opt = {
            k: 0
        };
        // let rows = [];
        let trLen = trs.length;
        // for (let i = 0; i < trLen; i++) {
        //     rows.push([]);
        // }
        // console.log(rows[0].length)
        for (let i = 0; i < trLen; i++) {
            let cells = getCellEls(trs[i]);
            opt.k = 0;
            for (let k = 0, cellsLen = cells.length; k < cellsLen; k++) {
                let cell = cells[k];
                // let { rowspan, colspan } = getCellElSpan(cell);
                // setTextToRows(cell, rows, i, k,opt);
                opt.k = opt.k > k ? opt.k : k;
                fillCellValue(cell, cellOpt, i, opt, config);
                opt.k++;
            }
            if (i === 0) {
                cellOpt.len.r = trLen;
                cellOpt.len.c = opt.k;
            }
            // rows.push(lineDatas);
        }
        opt = null;
        return cellOpt;
    } catch (error) {
        console.error(error);
    } finally {
        cellOpt = null;
    }
}
// function eachSpan(options,callback) {
//     let { rowspan = 1, colspan = 1 } = options;
//     // if()
// }
function fillCellValue(cell, cellOpt, rowIndex, opt, { startRow = 0, startCol = 0 } = { startRow: 0, startCol: 0 }) {
    let { rowspan, colspan } = getCellElSpan(cell);
    let text = getCellElText(cell);
    let { cellInfo, mergeInfo } = cellOpt;
    const mergeName = 'ly-merge';
    const mergeFcn = ({ sr, sc, er, ec }) => {
        mergeInfo.push({
            s: encodeAddress2(sr + startRow, sc + startCol),
            e: encodeAddress2(er + startRow, ec + startCol)
        });
    };
    const valueFcn = (r, c, tag = true) => {
        r = r + startRow;
        c = c + startCol;
        if (isEmpty(cellInfo[encodeAddress2(r, c)])) {
            if (tag) {
                cellInfo[encodeAddress2(r, c)] = text;
            } else {
                cellInfo[encodeAddress2(r, c)] = mergeName;
            }
        } else {
            opt.k++;
            while (!isEmpty(cellInfo[encodeAddress2(r, opt.k + startCol)])) {
                opt.k++;
            }
            cellInfo[encodeAddress2(r, opt.k + startCol)] = text;
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
}
// function fillCellValue1(cell, cellInfo, rowIndex, opt) {
//     let { rowspan, colspan } = getCellElSpan(cell);
//     let text = getCellElText(cell);
//     // console.log(encodeAddress2(rowIndex, opt.k))
//     const mergeName = 'ly-merge';
//     if (rowspan === 1 && colspan === 1) {
//         if (isEmpty(cellInfo[encodeAddress2(rowIndex, opt.k)])) {
//             cellInfo[encodeAddress2(rowIndex, opt.k)] = text;
//         } else {
//             opt.k++;
//             while (!isEmpty(cellInfo[encodeAddress2(rowIndex, opt.k)])) {
//                 opt.k++;
//             }
//             cellInfo[encodeAddress2(rowIndex, opt.k)] = text;
//         }
//     } else if (rowspan > 1 && colspan === 1) {
//         for (let r = rowIndex; r < rowIndex + rowspan; r++) {
//             if (isEmpty(cellInfo[encodeAddress2(r, opt.k)])) {
//                 if (r === rowIndex) {
//                     cellInfo[encodeAddress2(r, opt.k)] = text;
//                 } else {
//                     cellInfo[encodeAddress2(r, opt.k)] = mergeName;
//                 }
//             }
//         }
//     } else if (rowspan === 1 && colspan > 1) {
//         let { k } = opt;
//         for (let c = opt.k; c < opt.k + colspan; c++) {
//             if (isEmpty(cellInfo[encodeAddress2(rowIndex, c)])) {
//                 if (c === k) {
//                     cellInfo[encodeAddress2(rowIndex, c)] = text;
//                 } else {
//                     cellInfo[encodeAddress2(rowIndex, c)] = mergeName;
//                 }
//             } else {
//                 opt.k++;
//                 while (!isEmpty(cellInfo[encodeAddress2(rowIndex, opt.k)])) {
//                     opt.k++;
//                 }
//                 cellInfo[encodeAddress2(rowIndex, opt.k)] = text;
//             }
//         }
//     } else if (rowspan > 1 && colspan > 1) {
//         let { k } = opt;
//         for (let r = rowIndex; r < rowIndex + rowspan; r++) {
//             for (let c = opt.k; c < opt.k + colspan; c++) {
//                 if (isEmpty(cellInfo[encodeAddress2(r, c)])) {
//                     if (r === rowIndex && c === k) {
//                         cellInfo[encodeAddress2(r, c)] = text;
//                     } else {
//                         cellInfo[encodeAddress2(r, c)] = mergeName;
//                     }
//                 }
//             }
//         }
//     }
// }
// function fillRowsValue(cell, rows, rowIndex, opt) {
//     let { rowspan, colspan } = getCellElSpan(cell);
//     let text = getCellElText(cell);
//     const mergeName = 'ly-merge';
//     if (rowspan === 1 && colspan === 1) {
//         if (isEmpty(rows[rowIndex][opt.k])) {
//             rows[rowIndex][opt.k] = text;
//         } else {
//             opt.k++;
//             while (!isEmpty(rows[rowIndex][opt.k])) {
//                 opt.k++;
//             }
//             rows[rowIndex][opt.k] = text;
//         }
//     } else if (rowspan > 1 && colspan === 1) {
//         for (let r = rowIndex; r < rowIndex + rowspan; r++) {
//             if (isEmpty(rows[r][opt.k])) {
//                 if (r === rowIndex) {
//                     rows[r][opt.k] = text;
//                 } else {
//                     rows[r][opt.k] = mergeName;
//                 }
//             } else {
//                 // opt.k++;
//                 // while (!isEmpty(rows[rowIndex][opt.k])) {
//                 //     opt.k++;
//                 // }
//                 // rows[rowIndex][opt.k] = text;
//             }
//         }
//     } else if (rowspan === 1 && colspan > 1) {
//         let { k } = opt;
//         for (let c = opt.k; c < opt.k + colspan; c++) {
//             if (isEmpty(rows[rowIndex][c])) {
//                 if (c === k) {
//                     rows[rowIndex][c] = text;
//                 } else {
//                     rows[rowIndex][c] = mergeName;
//                 }
//             } else {
//                 opt.k++;
//                 while (!isEmpty(rows[rowIndex][opt.k])) {
//                     opt.k++;
//                 }
//                 rows[rowIndex][opt.k] = text;
//             }
//         }
//     } else if (rowspan > 1 && colspan > 1) {
//         let { k } = opt;
//         for (let r = rowIndex; r < rowIndex + rowspan; r++) {
//             for (let c = opt.k; c < opt.k + colspan; c++) {
//                 if (isEmpty(rows[r][c])) {
//                     if (r === rowIndex && c === k) {
//                         rows[r][c] = text;
//                     } else {
//                         rows[r][c] = mergeName;
//                     }
//                 } else {
//                     // opt.k++;
//                     // while (!isEmpty(rows[r][opt.k])) {
//                     //     opt.k++;
//                     // }
//                     // rows[rowIndex][opt.k] = text;
//                 }
//             }
//         }
//     }
// }

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