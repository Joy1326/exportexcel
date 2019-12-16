/* eslint-disable */
import { saveAs } from 'file-saver';
import exportExcelFcn from './export';
export { getBase64Image } from './image_util';
import { isObject } from './utils';
import { forWokerData } from './workerfmt';
export default function exportExcel(options, config = {}) {
    return new Promise((resolve, reject) => {
        if (!options) {
            console.error('options不能为空！');
            return reject(new Error('options不能为空！'));
        }
        if (!isObject(options)) {
            console.error('options必须是object！');
            return reject(new Error('options必须是object！'));
        }
        forWokerData(options);
        let { filename = "下载", suffixName = '.xlsx' } = options;
        exportExcelFcn(options, config).then(buffer => {
            toSaveAs(buffer, filename + suffixName);
            resolve();
        }).catch(error => {
            console.error(error);
            return reject(error);
        });
    })
}
export function exportExcelUseWorker(options, config = {}) {
    let _rootPath = '';
    let _url = './export-worker/export.js';
    if (window.workerFileRootPath) {
        _rootPath = window.workerFileRootPath;
    };
    if (window.workerFileUrl) {
        _url = window.workerFileUrl;
    }
    return new Promise((resolve, reject) => {
        let { filename = "下载", suffixName = '.xlsx' } = options;
        forWokerData(options, true);
        const wk = new Worker(_rootPath+_url);
        wk.postMessage({ options, config });
        wk.onmessage = function (e) {
            if(e.data===-1){
                reject(new Error('error'));
                return;
            }
            toSaveAs(e.data, filename + suffixName);
            resolve();
        }
        wk.onerror = function (error) {
            console.error(error);
            reject(error);
        }
    });
}
function toSaveAs(buffer, fileName) {
    try {
        saveAs(new Blob([buffer]), fileName);
    } catch (error) {
        console.error(error);
    }
}
export function canExport(showTip = true, msg = "该浏览器不支持前端导出功能，请升级浏览器！") {
    let canExp = true;
    if (typeof Worker === 'undefined') {
        canExp = false;
        if (showTip) {
            alert(msg);
        }
    }
    return canExp;
}