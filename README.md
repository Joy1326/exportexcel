# 基本用法
``` js
import exportExcel from './excel';

//导出
exportExcel(options,config);
```

# options
``` js
{
    table?:tableProps,
    tables?:tablesProps,
    images?:sheetImages,
    backgroundImage?:sheetBackgroundImage,
    filename?:'文件名',
    sheetname?:'工作表名',
    suffixName?:'文件后缀名',
    sheetProps?sheetProps,
}
```
# tableProps
```js
document.querySeletor('#table')
```
```js
{
    el?:document.querySeletor('#table'),
    space?:spaceProps,
    orgin?origProps,
    mergeCells?:({row,rowIndex,key,keyIndex})=>{rowspan?:number,colspan?:number},
    rowStyle?({row,rowIndex,key,keyIndex})=>style
}
```

# spaceProps
设置距离边
```js
{
    left:number,
    top:number,
    right:number,
    bottom:number
}
```
# originProps
定位，
```js
"A1"
//或者
{
    col:number,
    row:number
}
```