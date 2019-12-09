<template>
  <div>
    <h2>table</h2>
    <div class="card">
      <div>
        <input
          type="button"
          @click="tableExport(1)"
          value="table导出"
        >
        <input
          type="button"
          @click="tableExport(2)"
          value="table导出(自定义sheetname和filename)"
        >
        <input
          type="button"
          @click="tableExport(3)"
          value="table导出(对象配置方式)"
        >
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
    <h2>tables</h2>
    <div class="card">
      <input
        type="button"
        @click="tablesExport(1)"
        value="tables导出"
      >
    </div>
  </div>
</template>
<script>
import exportExcel, { readWorkbookFromRemoteFile } from "../lib/excel";
export default {
  methods: {
    getTable() {
      return this.$refs.table;
    },
    tableExport(type) {
      switch (type) {
        case 1:
          exportExcel({
            table: this.getTable()
          });
          break;
        case 2:
          exportExcel({
            table: this.getTable(),
            sheetname: "mysheet",
            filename: "测试"
          });
          break;
        case 3:
          exportExcel({
            table: {
              header: this.getHeader(),
              data: this.getData()
            }
          });
          break;
        default:
          break;
      }
    },
    tablesExport(type) {
      switch (type) {
        case 1:
          exportExcel({
            tables: [
              [
                this.getTable(),
                {
                  el: this.getTable(),
                  rowStyleList:[
{
                      index: 3,
                      style: {
                        font: {
                          color: { argb: "FFFF0000" }
                        }
                      }
                    }
                  ],
                  space: {
                    left: 3,
                    top: 2
                  }
                }
              ],
              [
                {
                  el: this.getTable(),
                  space: {
                    left: 2
                  },
                  mergeCellsList:[
                    {
                      rowIndex:2,
                      keyIndex:4,
                      rowspan:2
                    },
                    {
                      rowIndex:2,
                      keyIndex:0,
                      colspan:2
                    }
                  ],
                  rowStyleList: [
                    {
                      index: 3,
                      style: {
                        font: {
                          color: { argb: "FFFF0000" }
                        }
                      }
                    },
                    {
                      index: 6,
                      style: {
                        font: {
                          color: { argb: "FFFF0000" }
                        }
                      }
                    }
                  ]
                }
              ]
            ]
          });
          break;
      }
    },
    getHeader() {
      return [
        {
          title: "姓名",
          key: "name"
        },
        {
          title: "年龄",
          key: "age"
        },
        {
          title: "基本信息",
          children: [
            {
              title: "地址",
              key: "address"
            },
            {
              title: "电话",
              key: "tel"
            },
            {
              title: "联系",
              fmt: ({ row }) => {
                return row.address + row.tel;
              }
            }
          ]
        }
      ];
    },
    getData() {
      return [
        {
          name: "张三",
          age: 20,
          address: "张三地址",
          tel: "123456"
        },
        {
          name: "李四",
          age: 33,
          address: "李四地址",
          tel: "8888888"
        }
      ];
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