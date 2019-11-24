<template>
  <div>
    <div>
      <button @click="expFn(1)">dom导出</button>
      <button @click="expFn(2)">dom导出-no body</button>
      <button @click="expFn(3)">dom导出-no header</button>
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
import exportExcel, { tableToJson } from "../lib/excel/dom";
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
          console.log(el)
      exportExcel(el);
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
    getData() {
      let data = [];
      for (let i = 0; i < 10; i++) {
        let keyMap = {};
        for (let j = 0; j < 8; j++) {
          let key = String.fromCharCode(65 + j).toLowerCase();
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