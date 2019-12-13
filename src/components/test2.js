/* eslint-disable */
import exportExcel1, {
    exportExcelUseWorker as exportExcel,
    getBase64Image
} from "../export";
import logo from "../assets/logo.png";
import img from "../assets/a.jpg";
export default function testExport(type, table) {
    type = 10;
    switch (type) {
        case 1:
            export1(table);
            break;
        case 2:
            export2(table);
            break;
        case 3:
            export3(table);
            break;
        case 4:
            export4(table);
            break;
        case 5:
            export5(table);
            break;
        case 6:
            export6(table);
            break;
        case 7:
            export7(table);
            break;
        case 8:
            export8(table);
            break;
        case 9:
            export9(table);
            break;
        case 10:
            export10(table);
            break;
        default:
            break;
    }
}
// 基本使用
function export1(table) {
    let ff = { a: 'a' };
    let o = {
        table: table,
        gg() {
            console.log('dsfs')
            console.log(ff)
        }
    };
    exportExcel(o);
}
// 定位到单元格G3
function export2(table) {
    exportExcel({
        table: {
            el: table,
            origin: "G3"
        },
        filename: '从G3单元格开始'
    })
}
// 使用样式
function export3(table) {
    let o = {
        table: {
            el: table,
            rowStyle: rowStyle
        },
        filename: '样式'
    };
    exportExcel(o);
}
// 间隔
function export4(table) {
    exportExcel({
        table: {
            el: table,
            space: {
                top: 4,
                left: 1
            }
        },
        filename: '设置间隔'
    })
}
// 五个表格组成
function export5(table) {
    exportExcel({
        tables: [
            [table],
            [{
                el: table,
                space: {
                    top: 2,
                    bottom: 3
                }
            }],
            [{
                el: table,
                origin: 'N7',
                rowStyle: rowStyle
            }, {
                el: table,
                space: {
                    left: 1,
                    right: 2
                }
            }, table]
        ],
        filename: '多表格输出'
    })
}
// 多个sheets单个table
function export6(table) {
    exportExcel({
        sheets: [{
            table: table,
            sheetname: 'sheet1'
        }, {
            table: {
                el: table,
                origin: 'G2',
                rowStyle: rowStyle
            },
            sheetname: 'ss'
        }],
        filename: '多个sheets单个table'
    })
}
// 多个sheets多个tables
function export7(table) {
    exportExcel({
        sheets: [{
            tables: [
                [table, table],
                [table]
            ],
        }, {
            tables: [
                [{
                    el: table,
                    space: {
                        right: 1
                    }
                }, table]
            ]
        }],
        filename: '多个sheets多个tables'
    })
}
// 多个sheets多个tables及table
function export8(table) {
    exportExcel({
        sheets: [{
            tables: [
                [table, table],
                [table]
            ],
            gg(){}
        }, {
            table: table
        }, {
            table: {
                el: table,
                origin: 'D5'
            }
        }],
        filename: '多个sheets多个tables及table'
    })
}
// 添加图片
async function export9(table) {
    let bg = await getBase64Image(img);
    exportExcel({
        table: table,
        images: [{
            base64: bg,
            range: "B1:G8"
        }, {
            base64: bg,
            range: "L10:M11"
        }],
        filename: '添加单元格图片'
    });
}
// 添加背景图片及图片
async function export10(table) {
    let bg = await getBase64Image(img);
    let lg = await getBase64Image(logo);
    exportExcel({
        table: table,
        backgroundImage: {
            base64: bg
        },
        images: [{
            base64: lg,
            range: "B1:G8"
        }, {
            base64: lg,
            range: "L10:M11"
        }],
        filename: '背景图片-单元格图片'
    });
}
function rowStyle({ rowIndex }) {
    if ({ rowIndex }) {
        if (rowIndex === 2) {
            return {
                font: {
                    color: { argb: 'FFC107' }
                }
            }
        }
    }
}