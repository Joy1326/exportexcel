# 基本使用
## 安装依赖
```js
npm install exceljs
npm install file-saver
```
## 基本使用
```js
import exportExcel from './export';

exportExcel({
    table:this.$refs.table
});
```
# table
```js
{
    table:tableDOM || tableOptions,
    images?:imagesOptions,
    backgroundImage?:backgroundImageOptions,
    sheetname?:'sheet1',
    filename?:'下载'
}
```
# tables
```js
{
    tables:tablesOptions,
    images?:imagesOptions,
    backgroundImage?:backgroundImageOptions,
    sheetname?:'sheet1',
    filename?:'下载'
}
```
# sheets
```js
{
    sheets:[{
        table?:tableDOM|| tableOptions,
        tables?:tablesOptions,
        sheetname?:'sheet1'
    }],
    filename?:'下载'
}
```
# tableOptions
```js
{
    el?:tableDOM,
    header?:,
    keys?:,
    data?:[],
    rowStyle?:({row,rowIndex,key,keyIndex})=>styleOptions,
    mergeCells?:({row,rowIndex,key,keyIndex})=>{colspan?:number,rowspan?:number,value?:string||number||({row,rows,rowIndex,key,keys,keyIndex})},
    space?:{left?:number,top?:number,right?:number,bottom?:number},//表格间隔
    origin?:'A1'||{col:number,row:number} // 定位到单元格
}
```
# tablesOptions
```js
[
    [tableDOM||tableOptions,tableDOM||tableOptions,...],
    [...],
    ...
]

```
# imagesOptions
```js
[{
    base64:'base64图片文件',
    range:string,// 例如：'B1:C8',
    extension?:'png'
}]
```
# backgroundImageOptions
```js
{
    base64:'base64图片文件',
    extension?:'png'
}
```

