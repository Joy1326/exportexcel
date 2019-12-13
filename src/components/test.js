/* eslint-disable */
import exportExcel, {
    // exportExcelUseWorker as exportExcel,
    getBase64Image
} from "../export";
import logo from "../assets/logo.png";
import img from "../assets/a.jpg";
export default function testExport(type, table) {
    type = 3;
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
    exportExcel({
        table: {
            header: getHead(),
            data: getData()
        },
    });
}
// 定位到单元格G3
function export2(table) {
    exportExcel({
        table: {
            header: getHead(),
            data: getData(),
            origin: "G3"
        },
        filename: '从G3单元格开始'
    })
}
// 使用样式,合并
function export3(table) {
    exportExcel({
        table: {
            header: getHead(),
            data: getData(),
            rowStyle: rowStyle,
            mergeCells({ rowIndex, key }) {
                if (rowIndex === 1&&key==='G') {
                    return {
                        colspan:2
                    }
                }
                if (rowIndex === 3 && key === "A") {
                    return {
                        rowspan:3
                    }
                }
                if (rowIndex === 3 && key === "D") {
                    return {
                        rowspan: 2,
                        colspan: 3,
                        // value:'SFSF'
                        value({rows,row,key}) {
                            return row[key]+'___'+rows[0][key]+row.A
                        }
                    }
                }
            }
        },
        filename: '样式-合并'
    })
}
// 间隔
function export4(table) {
    exportExcel({
        table: {
            header: getHead(),
            data: getData(),
            space: {
                top: 4,
                left: 1
            }
        },
        filename: '设置间隔'
    })
}
// 多表格组成
function export5(table) {
    exportExcel({
        tables: [
            [{
                header: getHead(),
                data: getData()
            }],
            [{
                header: getHead(),
                data: getData()
            }]
        ],
        filename: '多表格输出'
    })
}
// 多个sheets单个table
function export6(table) {
    exportExcel({
        sheets: [{
            table: {
                header: getHead(),
                data: getData()
            },
            sheetname: 'sheet1'
        }, {
            table: {
                header: getHead(),
                data: getData(),
                origin: 'G2',
                rowStyle: rowStyle
            },
            sheetname: 'ss'
        }],
        filename: '多个sheets单个table'
    })
}
// 添加图片
async function export7(table) {
    let bg = await getBase64Image(img);
    exportExcel({
        table: {
            header: getHead(),
            data: getData()
        },
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
async function export8(table) {
    let bg = await getBase64Image(img);
    let lg = await getBase64Image(logo);
    exportExcel({
        table: {
            header: getHead(),
            data: getData()
        },
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
// 读取xlsx文件
function export9() {
    exportExcel({
        xlsxFile: './b.xlsx',
        filename: '读取excel导出'
    })
}
// 读取xlsx文件,并追加内容，sheetname必须和xlsx中sheetname一样，否则新添加sheet工作表
function export10() {
    exportExcel({
        xlsxFile: './b.xlsx',
        table: {
            header: getHead(),
            origin:'F9'
        },
        sheetname: '你好',
        filename: '读取excel导出并追加新table'
    })
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
function getHead() {
    return [
        {
            key: "A",
            title: "A-title"
        },
        {
            key: "B",
            title: "B-title",
            children: [
                {
                    key: "G",
                    title: "G-title"
                },
                {
                    key: "H",
                    title: "H-title"
                }
            ]
        },
        {
            key: "C",
            title: "C-title",
            fmt({ row, key }) {
                return row.D + "0000";
            }
        },
        {
            key: "D",
            title: "D-title"
        },
        {
            key: "E",
            title: "E-title",
            children: [
                {
                    key: "I",
                    title: "I-title",
                    cellStyle({ rowIndex }) {
                        if (rowIndex === 2) {
                            return {
                                font: {
                                    color: { argb: "FFFF0000" }
                                }
                            };
                        }
                        if (rowIndex === 6) {
                            return {
                                font: {
                                    color: { argb: "FFFF0000" }
                                }
                            };
                        }
                    }
                },
                {
                    key: "J",
                    title: "J-title",
                    children: [
                        {
                            key: "K",
                            title: "K-title"
                        }
                    ]
                },
                {
                    key: "L",
                    title: "L-title"
                }
            ]
        },
        {
            key: "F",
            title: "F-title"
        }
    ];
}
function getData() {
    let data = [];
    for (let i = 0; i < 10; i++) {
        let keyMap = {};
        for (let j = 0; j < 18; j++) {
            let key = String.fromCharCode(65 + j).toUpperCase();
            keyMap[key] = key + i + "-value";
        }
        data.push(keyMap);
    }
    // console.log(JSON.stringify(data))
    return data;
}