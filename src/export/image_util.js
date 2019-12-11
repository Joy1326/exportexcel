function getUrlBase64(img, extension, quality = 1) {
    let canvas = document.createElement("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    let ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, img.width, img.height);
    let ext = extension || img.src.substring(img.src.lastIndexOf(".") + 1).toLowerCase();
    // toDataURL方法，可以是image/jpeg或image/webp,默认image/png
    let dataURL = canvas.toDataURL("image/" + ext, quality);
    return dataURL;
}
async function loadImage(src) {
    return new Promise((resolve, reject) => {
        let img = new Image();
        img.src = src;
        img.onload = function () {
            resolve(img);
        };
        img.onerror = function (error) {
            reject(error);
        };
    });
}
async function getBase64Image(imgSrc, extension = '', quality = 1) {
    let img = await loadImage(imgSrc);
    return getUrlBase64(img, extension, quality);
}
export { getBase64Image };