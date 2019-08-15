function typeOf(obj) {
    const toString = Object.prototype.toString;
    const map = {
        '[object Boolean]': 'boolean',
        '[object Number]': 'number',
        '[object String]': 'string',
        '[object Function]': 'function',
        '[object Array]': 'array',
        '[object Date]': 'date',
        '[object RegExp]': 'regExp',
        '[object Undefined]': 'undefined',
        '[object Null]': 'null',
        '[object Object]': 'object'
    };
    return map[toString.call(obj)];
}
// deepCopy
function deepCopy(data) {
    const t = typeOf(data);
    let o;

    if (t === 'array') {
        o = [];
    } else if (t === 'object') {
        o = {};
    } else {
        return data;
    }

    if (t === 'array') {
        for (let i = 0; i < data.length; i++) {
            o.push(deepCopy(data[i]));
        }
    } else if (t === 'object') {
        for (let i in data) {
            o[i] = deepCopy(data[i]);
        }
    }
    return o;
}

// set forTableHead to true when convertToRows, false in normal cases like table.vue
const getAllColumns = (cols, forTableHead = false) => {
    const columns = deepCopy(cols);
    const result = [];
    columns.forEach((column) => {
        if (column.children) {
            if (forTableHead) { result.push(column); }
            result.push.apply(result, getAllColumns(column.children, forTableHead));
        } else {
            result.push(column);
        }
    });
    return result;
};


const convertToRows = (columns) => {
    const originColumns = deepCopy(columns);
    let maxLevel = 1;
    const traverse = (column, parent) => {
        if (parent) {
            column.level = parent.level + 1;
            if (maxLevel < column.level) {
                maxLevel = column.level;
            }
        }

        if (column.children) {
            let colSpan = 0;
            column.children.forEach((subColumn) => {
                traverse(subColumn, column);
                colSpan += subColumn.colSpan;
            });
            column.colSpan = colSpan;
        } else {
            column.colSpan = 1;
        }
    };

    originColumns.forEach((column) => {
        column.level = 1;
        traverse(column);
    });

    const rows = [];
    for (let i = 0; i < maxLevel; i++) {
        rows.push([]);
    }

    const allColumns = getAllColumns(originColumns, true);
    let merges = [];
    let cell = [];
    let INDEX = 0;
    allColumns.forEach((column) => {
        let level = column.level;
        if (!column.children) {
            column.rowSpan = maxLevel - level + 1;
        }
        let rowSpan = column.rowSpan && column.rowSpan > 1 ? column.rowSpan - 1 : 0;
        let colSpan = column.colSpan && column.colSpan > 1 ? column.colSpan - 1 : 0;
        merges.push({
            s: {
                r: level - 1,
                c: INDEX
            },
            e: {
                r: level - 1 + rowSpan,
                c: INDEX + colSpan
            },
        });
        cell.push({
            col: {
                r: level - 1,
                c: INDEX,
            },
            key:column.key,
            title: column.title,
            colSpan: colSpan,
            rowSpan:rowSpan
        });
        if (column.children) {
            column.rowSpan = 1;
            INDEX--;
        }
        rows[level - 1].push(column);
        INDEX++;
    });
    return {
        rowsLen: rows.length,
        cell,
        merges,
        cellLen: INDEX,
    };
};

