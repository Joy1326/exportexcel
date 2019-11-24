import colCache from 'exceljs/lib/utils/col-cache';
export { colCache };
export function isDOM(str) {
    return str instanceof HTMLElement;
}
export function isObject(str) {
    return str instanceof Object;
}
export function isArray(str) {
    return str instanceof Array;
}
export function isEmpty(val) {
    return val === undefined || val === null;
}