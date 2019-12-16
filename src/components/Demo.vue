<template>
  <div>
    <div>
      <input
        type="button"
        value="导出"
        @click="exportFcn1"
      />
    </div>
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
          <td
            :data-cell-style1="cellStyle"
            rowspan="2"
          >6序号</td>
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
          <td :data-cell-style1="cellStyle1">5序号</td>
          <td>6序号</td>
          <td>7序号</td>
          <td>8序号</td>
          <td>9序号</td>
        </tr>
        <tr :data-row-style1="rowStyle">
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
</template>
<script>
/* eslint-disable no-console */
/* eslint-disable */
// import exportExcel, {
//   readWorkbookFromRemoteFile,
//   tableToJson,
//   exportExcelUseWorker,
//   getBase64Image
// } from "../export";
import testExport from './test';
// import img from "../assets/logo.png";
// import img1 from "../assets/a.jpg";
export default {
  data() {
    return {
      rowStyle: JSON.stringify({
        font: {
          color: { argb: "ff5722" }
        }
      }),
      cellStyle: JSON.stringify({
        font: {
          color: { argb: "FF00FF00" },
          bold: true
        }
      }),
      cellStyle1: JSON.stringify({
        fill: {
          type: "pattern",
          pattern: "darkTrellis",
          fgColor: { argb: "FFFFFF00" },
          bgColor: { argb: "FF0000FF" }
        }
      })
    };
  },
  methods: {
    getRefTable() {
      return this.$refs.table;
    },
    exportFcn1(){
        testExport(5,this.getRefTable());

      // for(let i=1;i<=10;i++){
      //   testExport(i,this.getRefTable());
      // }
    },
    async export1() {
      // let a = tableToJson(this.getRefTable())
      // console.log(a)
      // let wk = await readWorkbookFromRemoteFile("./b1.xlsx");
      // console.log(wk)
      // exportExcelUseWorker({
        // sheets:[{
        //   tables:[
        //     [this.getRefTable(),{el:this.getRefTable()}],
        //     [{
        //       el:this.getRefTable(),
        //       rowStyle({ rowIndex }) {
        //         if (rowIndex === 2) {
        //           return {
        //             font: {
        //               color: { argb: "2196F3" }
        //             }
        //           };
        //         }
        //       }
        //       },this.getRefTable()]
        //   ]
        // }],
        // sheets:[{
        //   table:{
        //     el:this.getRefTable(),
        //   }
        // }],
        // sheets:[{
        //   table:this.getRefTable()
        // }]
        // tables: [
        //   [this.getRefTable()],
        //   [
        //     {
        //       el: this.getRefTable(),
        //       rowStyle({ rowIndex }) {
        //         if (rowIndex === 2) {
        //           return {
        //             font: {
        //               color: { argb: "2196F3" }
        //             }
        //           };
        //         }
        //       }
        //     },
        //     this.getRefTable()
        //   ]
        // ],
        // table:{
        //   el:this.getRefTable(),
        //   rowStyle({rowIndex}){
        //     if(rowIndex===2){
        //       return {
        //         font:{
        //           color: { argb: "2196F3" }
        //         }
        //       }
        //     }
        //   },
        //   gg:2
        // },
        // table:this.getRefTable()
        // table:{
        //   header:this.getHead2()
        // }
      // });
      // let bg = await getBase64Image(img1);

      // exportExcel({
      //   sheets: [
      //     {
      //       table: this.getRefTable(),
      //       backgroundImage:{
      //         base64:await getBase64Image(img)
      //       }
      //     },
      //     {
      //       table: this.getRefTable(),
      //       backgroundImage:{
      //         base64:bg
      //       }
      //       images: [
      //         {
      //           base64: bg,
      //           range: "B1:G8"
      //         },
      //         {
      //           base64: bg,
      //           range: "L10:M11"
      //         }
      //       ]
      //     }
      //   ]
      // });
      // exportExcel({
      //   table:this.getRefTable(),
      //   backgroundImage:{
      //     base64:bg
      //   }
      // });
      // exportExcel({
      //   table:this.getRefTable(),
      //   images:[{
      //     base64:await getBase64Image(img),
      //     range:'B1:G8'
      //   },{
      //     base64:await getBase64Image(img),
      //     range:'L10:M11'
      //   }]
      // });
      // exportExcel({
      //   xlsxFile:'./b.xlsx',
      //   table:{
      //     el:this.getRefTable(),
      //     origin:'K3'
      //   },
      //   sheetname:'你好'
      // })
      // exportExcel({
      //   tables:[
      //     [this.getRefTable(),{
      //       el:this.getRefTable(),
      //       space:{
      //         left:2,
      //         top:3
      //       }
      //     }],[{
      //       header:this.getHead2(),
      //       data:this.getData()
      //     },this.getRefTable()]
      //   ],
      //   sheetname:'f'
      // })
      // exportExcel({
      //   table:{
      //     header:this.getHead2(),
      //     data:this.getData(),
      //     origin:'E6'
      //   },
      //   sheetname:'gg',
      //   filename:'tt'
      // });
      // exportExcel({
      //   table:{
      //     el:this.getRefTable(),
      //     origin:'E3'
      //   },
      //   sheetname:'gg',
      //   filename:'tt'
      // });
      // exportExcel({
      //   table:this.getRefTable(),
      //   sheetname:'gg',
      //   filename:'tt'
      // });
      // exportExcel({
      //   xlsxFile:'./b.xlsx'
      // })
      // console.log(wk);
      // exportExcelUseWorker(wk);
      // exportExcel({
      //   sheets: [
      //     {
      //       table: {
      //         el: this.getRefTable(),
      //         rowStyle({ row, rowIndex, key, keyIndex }) {
      //           // console.log(row, rowIndex, key, keyIndex);
      //           if (rowIndex === 2) {
      //             return {
      //               font: {
      //                 color: { argb: "2196F3" }
      //               }
      //             };
      //           }
      //         }
      //       }
      //     }
      //   ]
      // });
      // exportExcel({
      //   sheets:[{
      //     table:{
      //       header:this.getHead2(),
      //       // keys:['A',"C"],
      //       data:this.getData()
      //     }
      //   }]
      // })
      // exportExcel({
      //   sheets:[{
      //     table:this.getRefTable()
      //   }],
      //   filename:'测试'
      // });
      // exportExcel({
      //   sheets:[{
      //     table:{
      //       el:this.getRefTable()
      //     }
      //   }],
      //   filename:'测试'
      // });
      // exportExcel({
      //   sheets:[{
      //     tables:[
      //       [this.getRefTable(),{
      //         el:this.getRefTable(),
      //         space:{
      //           left:2
      //         }
      //       }],
      //       [{
      //         el:this.getRefTable(),
      //         space:{
      //           top:3,
      //           left:4
      //         }
      //         }],
      //       [{
      //         el:this.getRefTable(),
      //         origin:'Q10',
      //         }],
      //       [{
      //         el:this.getRefTable(),
      //         space:{
      //           top:1
      //         }
      //       }]
      //     ]
      //   }],
      //   filename:'测试'
      // });
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
              gg() {},
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
.card {
  padding: 20px;
  margin: 6px;
  border: 1px solid #cac9c9;
  box-shadow: 0 0 5px 5px #c3cbce;
  border-radius: 5px;
}
table {
  width: 100%;
  margin: 5px 0;
}
table,
table td,
table th {
  border: 1px solid gray;
  border-collapse: collapse;
}
</style>