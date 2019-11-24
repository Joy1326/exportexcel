# 基本用法
```
exportExcel(options,config?);
```
---
## options配置
```
tableElement
```
table或者tables必须配置一个
```
{
    table?:tableOptions,
    tables?:tablesOptions,//如果是多个表格
    sheetname?:string, // default:Sheet1
    filename?:stirng, // default:下载
    suffixName?:string, // default:.xlsx
    sheetOptions?:sheetOptions
}
```
## tableOptions配置
```
tableElement
```
对象方式配置
```
{
    el:tableElement
}
```
```
{
    keys?:[],// 如果不需要显示头部（即不配置columns）,配置显示列的数据keyName
    header?:headerOptions,// 头部,不传时不显示头部，如果不传则必须传keys
    data:[], // [['value1','value2']] 或者 [{key1:'value1',key2:'value2'}]
    rowStyle?:({row,rowIndex,key,keyIndex})=>styleOptions
    mergeCell?:({row,rowIndex,key,keyIndex})=>{rowspan?,colspan?}||[{s:{r,c},e:{r,c}}]
}
```
## talbesOptions配置
```
[
    [tableOptions对象方式配置,...],
    [tableOptions对象方式配置,...]
]
```
## headerOptions配置
```
[{
    key:'keyName',
    title:'title',
    fmt:({row,rowIndex,key,keyIndex}),
    cellStyle:({row,rowIndex,key,keyIndex})=>styleOptions
    children:[]
}]
```
## sheetOptions
```
//TODO
```
## styleOptions
```
//TODO
```
## config配置（可选）
```
{
    // TODO
    workbookOptions
}
```
