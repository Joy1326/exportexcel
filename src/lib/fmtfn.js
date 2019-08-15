export function fmtFn(opt) {
    let { sheets } = opt;
    sheets.forEach(sheet => {
        let { tables } = sheet;
        fmtTableFnToString(tables);
    });
}
function fmtTableFnToString(tables) {
    tables.forEach(columnTables => {
        columnTables.forEach(table => {
            let { mergeCells, rowStyle } = table;
            if (mergeCells && typeof mergeCells === 'function') {
                table.mergeCells = mergeCells.toString();
            }
            if (rowStyle && typeof rowStyle === 'function') {
                table.rowStyle = rowStyle.toString();
            }
            fmtColumnFnToString(table);
        });
    });
}
function fmtColumnFnToString(table) {
    let { columns } = table;
    columns.forEach(column => {
        let { fmt, cellStyle } = column;
        if (fmt && typeof fmt === 'function') {
            column.fmt = fmt.toString();
        }
        if (cellStyle && typeof cellStyle === 'function') {
            column.cellStyle = cellStyle.toString();
        }
    });
}