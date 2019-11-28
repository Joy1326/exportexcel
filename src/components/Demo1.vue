<template>
  <div>
    <div>
      <button @click="expFn(1)">dom导出</button>
      <button @click="expFn(2)">dom导出-no body</button>
      <button @click="expFn(3)">dom导出-no header</button>
      <button @click="expFnObj">dom导出tableObj</button>
      <button @click="expFnObjMergeHeader">dom导出tableObjMergeHeader</button>
      <button @click="expFnSheets">dom导出expFnSheets</button>
      <button @click="expFnSheetsFromExcel">expFnSheetsFromExcel</button>
    </div>
    <div>
      <table ref="table">
        <thead>
          <tr>
            <th rowspan="3">0</th>
            <th>1</th>
            <th>2</th>
            <th colspan="2">3</th>
            <th>4</th>
            <th rowspan="3">5</th>
            <th>6</th>
            <th
              rowspan="2"
              colspan="2"
            >7</th>
            <th>8</th>
          </tr>
          <tr>
            <th>2-0</th>
            <th>2-1</th>
            <th>2-2</th>
            <th>2-3</th>
            <th>2-4</th>
            <th>2-5</th>
            <th>2-6</th>
          </tr>
          <tr>
            <th>3-0</th>
            <th>3-1</th>
            <th>3-2</th>
            <th colspan="2">3-3</th>
            <th>3-4</th>
            <th>3-5</th>
            <th>3-6</th>
            <th>3-7</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>1序号</td>
            <td>2序号</td>
            <td>3序号</td>
            <td>4序号</td>
            <td colspan="3">5序号</td>
            <td>6序号</td>
            <td>7序号</td>
            <td>8序号</td>
            <td>9序号</td>
          </tr>
          <tr>
            <td>1序号</td>
            <td colspan="2">2序号</td>
            <td>3序号</td>
            <td>4序号</td>
            <td>5序号</td>
            <td rowspan="2">6序号</td>
            <td>7序号</td>
            <td
              colspan="2"
              rowspan="3"
            >8序号</td>
            <td>9序号</td>
          </tr>
          <tr>
            <td>1序号</td>
            <td>2序号</td>
            <td>3序号</td>
            <td>4序号</td>
            <td>5序号</td>
            <td>7序号</td>
            <td>8序号</td>
            <td>9序号</td>
          </tr>
          <tr>
            <td>1序号</td>
            <td>2序号</td>
            <td>3序号</td>
            <td>4序号</td>
            <td>5序号</td>
            <td>6序号</td>
            <td>7序号</td>
            <td>8序号</td>
            <td>9序号</td>
          </tr>
          <tr>
            <td>1序号</td>
            <td>2序号</td>
            <td>3序号</td>
            <td>4序号</td>
            <td>5序号</td>
            <td>6序号</td>
            <td>7序号</td>
            <td>8序号</td>
            <td>9序号</td>
            <td>10序号</td>
            <td>11序号</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>
<script>
import exportExcel, {
  tableToJson,
  readWorkbookFromRemoteFile
} from "../lib/excel/export";
export default {
  data() {
    let data = this.getData();
    return {
      header: data.slice(0, 1),
      body: data.slice(1)
    };
  },
  methods: {
    expFn(index) {
      //   let datas = tableToJson(this.$refs.table);
      //   console.log(datas)
      // exportExcel(this.getSheets());
      let table = this.getRefTable();
      let el =
        index === 1
          ? table
          : index === 2
          ? table.querySelector("thead")
          : table.querySelector("tbody");
      console.log(el);
      exportExcel(el);
    },
    expFnObj() {
      exportExcel({
        table: {
          header: this.getHead1()
        }
      });
    },
    expFnObjMergeHeader() {
      exportExcel({
        table: {
          header: this.getHead2(),
          data: this.getData(),
          mergeCells({ rowIndex, key }) {
            if (rowIndex === 3 && key === "G") {
              return {
                colspan: 3
              };
            }
            if (rowIndex === 4 && key === "C") {
              return {
                rowspan: 2
              };
            }
            if (rowIndex === 2 && key === "I") {
              return {
                colspan: 2,
                rowspan: 3
              };
            }
          }
        }
      });
    },
    expFnSheets() {
      exportExcel({
        sheets: [
          {
            tables: [
              [
                this.getRefTable(),
                {
                  header: this.getHead2(),
                  data: this.getData(),
                  space: {
                    bottom: 2
                  },
                  origin: {
                    col: 14,
                    row: 2
                  }
                  // mergeCells({ rowIndex, key }) {
                  //   if (rowIndex === 3 && key === "G") {
                  //     return {
                  //       colspan: 3
                  //     };
                  //   }
                  //   if (rowIndex === 4 && key === "C") {
                  //     return {
                  //       rowspan: 2
                  //     };
                  //   }
                  //   if (rowIndex === 2 && key === "I") {
                  //     return {
                  //       colspan: 2,
                  //       rowspan: 3
                  //     };
                  //   }
                  // }
                }
              ],
              [
                this.getRefTable(),
                {
                  header: this.getHead2(),
                  data: this.getData(),
                  space: {
                    right: 3
                  },
                  origin: "M27",
                  rowStyle({ rowIndex, key }) {
                    if (rowIndex === 2) {
                      return {
                        font: {
                          name: "Arial Black",
                          color: { argb: "FF00FF00" },
                          family: 2,
                          size: 14,
                          italic: true
                        }
                      };
                    }
                    if (rowIndex === 4) {
                      return {
                        fill: {
                          type: "pattern",
                          pattern: "darkVertical",
                          fgColor: { argb: "FFFF0000" }
                        }
                      };
                    }
                  }
                },
                this.getRefTable()
              ]
            ],
            sheetname: "Sheet1"
          }
        ]
      });
    },
    async expFnSheetsFromExcel() {
      let wk = await readWorkbookFromRemoteFile("./a.xlsx");
      if (!wk) {
        return;
      }
      exportExcel(
        {
          sheets: [
            {
              tables: [[this.getRefTable()]]
            }
          ]
        },
        { workbook: wk }
      );
    },
    getRefTable() {
      return this.$refs.table;
    },
    getSheets() {
      return {
        sheets: [
          {
            tables: {},
            //{el:el,sheetname:''},{columns:[],data:[]},{keys:[],data:[],sheetname:''}
            table: {
              el: this.getRefTable()
            },
            sheetname: "Sheet1"
          }
        ],
        filename: "下载"
      };
    },
    getHead1() {
      return [
        {
          key: "A",
          title: "A-title"
        },
        {
          key: "B",
          title: "B-title"
        },
        {
          key: "C",
          title: "C-title"
        },
        {
          key: "D",
          title: "D-title"
        },
        {
          key: "E",
          title: "E-title"
        },
        {
          key: "F",
          title: "F-title"
        }
      ];
    },
    getHead2() {
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
          title: "C-title"
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
              gg(){},
              cellStyle({ rowIndex}) {
                if (rowIndex === 2) {
                  return {
                    font: {
                      color: { argb: "FFFF0000" }
                    }
                  };
                }
                if(rowIndex===6){
                  return {
                    font: {
                      color: { argb: "FFFF0000" }
                    }
                  }
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
    },
    getData() {
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
  }
};
</script>
<style scoped>
table {
  width: 100%;
}
table,
table td,
table th {
  border: 1px solid gray;
  border-collapse: collapse;
}
</style>