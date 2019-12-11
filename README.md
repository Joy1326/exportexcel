# sheets
```js
{
    sheets:[{
        table?:tableDOM|| tableOptions,
        tables?:[[tableDOM||tableOptions]],
        ws?:ws,//TODO
        sheetname:'sheet1'
    }],
    filename?:'下载'
}
```
# table
```js
{
    table:tableDOM || tableOptions,
    sheetname:'sheet1',
    filename:'下载'
}
```
# tables
```js
{
    tables:[[tableDOM||tableOptions]],
    sheetname:'sheet1',
    filename:'下载'
}
```
# tableOptions
```js
{
    el?:tableDOM,
    header?:,
    keys?:,
    data?:[],
    rowStyle?:Function,
    rowStyleList?:[],
    mergeCells?:Function,
    mergeCellsList?:[],
    space?:{left,top,right,bottom},
    origin?:'A1'
}
```

