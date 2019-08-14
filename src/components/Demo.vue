<template>
  <div class="hello">
    <button @click="expFn">导出</button>
  </div>
</template>

<script>
import exportExcell, { getBase64Image } from "../lib/export";
import list from "./data";
import img from "../assets/bg.jpg";
import img1 from "../assets/logo.png";
export default {
  name: "HelloWorld",
  methods: {
    async expFn() {
      let base64Img = await getBase64Image(img, "jpeg", 0.8);
      let wsbase64Img = await getBase64Image(img1, "jpeg", 0.8);
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
            backgroundImage: {
              base64: base64Img,
              extension: "jpeg"
            },
            wsImages: [
              {
                base64: wsbase64Img,
                range: "B2:D6"
              },
              {
                base64: wsbase64Img,
                range: {
                  tl: { col: 9.5, row: 1.5 },
                  br: { col: 12.5, row: 5.5 }
                }
              },
              {
                base64: wsbase64Img,
                range: {
                  tl: { col: 9.5, row: 10.5 },
                  br: { col: 12.5, row: 15.5 },
                  editAs: "oneCell"
                }
              }
            ],
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
            wsImages: [
              {
                base64: wsbase64Img,
                range: {
                  tl: { col: 0, row: 0 },
                  br:{col:4,row:4},
                  ext: { width: 500, height: 200 }
                }
              },
              {
                base64: wsbase64Img,
                range: {
                  tl: { col: 10, row: 0 },
                  br:{col:14,row:6},
                  editAs:'absolute'
                }
              }
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
