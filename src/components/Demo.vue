<template>
  <div class="hello">
    <button @click="expFn">导出</button>
  </div>
</template>

<script>
import exportExcell from "../lib/export";
import list from "./data";
export default {
  name: "HelloWorld",
  methods: {
    expFn() {
      exportExcell({
        sheets: [
          {
            tables: [
              [
                {
                  columns: this.getColumns("第一个"),
                  data: list,
                  space: {
                    left: 2,
                    top: 3,
                    bottom: 4,
                    right: 1
                  },
                  origin: {
                    r: 5,
                    c: 19
                  },
                  mergeCells: params => {
                    let { key, keyIndex, row, rowIndex } = params;
                    if (rowIndex === 8 && key === "C") {
                      // console.log(params)
                      return {
                        rowspan: 3,
                        colspan: 5
                      };
                    }
                  }
                },
                {
                  columns: this.getColumns2(),
                  data: list,
                  space: {
                    left: 1,
                    right: 1
                  },
                  mergeCells: params => {
                    let { key, keyIndex, row, rowIndex } = params;
                    if (rowIndex === 2 && key === "A") {
                      // console.log(params)
                      return {
                        rowspan: 2,
                        colspan: 3
                      };
                    }
                    if (rowIndex === 5 && key === "B") {
                      return {
                        colspan: 2
                      };
                    }
                  }
                },
                {
                  columns: this.getColumns("第三个"),
                  data: list
                }
              ],
              [
                {
                  columns: this.getColumns("第四个"),
                  data: list,
                  origin: {
                    c: 29
                  }
                },
                {
                  columns: this.getColumns2(),
                  data: list,
                  mergeCells: params => {
                    let { key, keyIndex, row, rowIndex } = params;
                    if (rowIndex === 2 && key === "B") {
                      // console.log(params)
                      return {
                        rowspan: 3,
                        colspan: 5
                      };
                    }
                  }
                },
                {
                  columns: this.getColumns("第五个"),
                  data: list
                }
              ]
            ],
            props: {
              properties: {
                tabColor: { argb: "FFC0000" }
              },
              views: [
                {
                  state: "frozen",
                  xSplit: 2,
                  ySplit: 3
                }
              ]
            },
            sheetName: "Sheet1"
          },
          {
            tables: [
              [
                {
                  columns: this.getColumns(),
                  data: list,
                  origin: {
                    r: 24,
                    c: 14
                  }
                },
                {
                  columns: this.getColumns2(),
                  data: list
                },
                {
                  columns: this.getColumns(),
                  data: list,
                  mergeCells: params => {
                    let { key, keyIndex, row, rowIndex } = params;
                    if (rowIndex === 2 && key === "F") {
                      // console.log(params)
                      return {
                        colspan: 3
                      };
                    }
                  }
                }
              ],
              [
                {
                  columns: this.getColumns(),
                  data: list
                },
                {
                  columns: this.getColumns2(),
                  data: list,
                  origin: {
                    r: 3,
                    c: 22
                  },
                  mergeCells: params => {
                    let { key, keyIndex, row, rowIndex } = params;
                    if (rowIndex === 8 && key === "D") {
                      // console.log(params)
                      return {
                        rowspan: 3,
                        colspan: 3
                      };
                    }
                  }
                },
                {
                  columns: this.getColumns(),
                  data: list
                }
              ]
            ],
            sheetName: "Sheet2"
          }
        ],
        fileName: "文件名"
      });
    },
    getColumns(num = "gg") {
      let obj = {
        num: num
      };
      return [
        {
          type: "index",
          title: "序号"
        },
        {
          title: "标题A",
          key: "A",
          params: obj,
          cellStyle: opt => {
            let { rowIndex } = opt;
            if (rowIndex === 1) {
              return {
                alignment: {
                  textRotation: 30,
                  vertical: "middle",
                  horizontal: "center"
                }
              };
            } else if (rowIndex === 8) {
              return {
                border: {
                  top: { style: "double", color: { argb: "FF00FF00" } },
                  left: { style: "double", color: { argb: "FF00FF00" } },
                  bottom: { style: "double", color: { argb: "FF00FF00" } },
                  right: { style: "double", color: { argb: "FF00FF00" } }
                }
              };
            }
          },
          fmt: opt => {
            // console.log(opt)
            let { column, row } = opt;
            return column.params.num + "_" + row.B;
          }
        },
        {
          title: "标题B",
          key: "B",
          cellStyle: opt => {
            if (opt.rowIndex === 2) {
              return {
                font: {
                  name: "Comic Sans MS",
                  family: 4,
                  size: 16,
                  underline: true,
                  bold: true
                }
              };
            } else if (opt.rowIndex === 5) {
              return {
                font: {
                  color: { argb: "FF00FF00" }
                }
              };
            } else if (opt.rowIndex === 8) {
              return {
                fill: {
                  type: "pattern",
                  pattern: "darkTrellis",
                  fgColor: { argb: "FFFFFF00" },
                  bgColor: { argb: "FF0000FF" }
                }
              };
            }
          }
        },
        {
          title: "标题C",
          key: "C"
        },
        {
          title: "标题D",
          children: [
            {
              title: "标题E",
              key: "E"
            },
            {
              title: "标题F",
              key: "F"
            }
          ]
        }
      ];
    },
    getColumns2() {
      return [
        {
          title: "序号",
          key: "index"
        },
        {
          title: "标题A",
          key: "A",
          fmt: opt => {
            let { row } = opt;
            return {
              text: row.A,
              hyperlink: "http://www.baidu.com",
              tooltip: "www.baidu.com"
            };
          }
        },
        {
          title: "标题B",
          key: "B"
        },
        {
          title: "标题D",
          key: "D"
        },
        {
          title: "标题E",
          children: [
            {
              title: "标题F",
              key: "F"
            },
            {
              title: "标题G",
              children: [
                {
                  title: "标题G",
                  key: "G"
                },
                {
                  title: "标题H",
                  key: "H"
                }
              ]
            }
          ]
        }
      ];
    }
  }
};
</script>
