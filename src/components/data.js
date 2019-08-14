let data = [];
for (let i = 0; i < 20; i++){
    let m = {};
    for (let key = 0; key < 15; key++){
        let k = String.fromCharCode(key + 65);
        m[k] = k;
        m['index'] = i + 1;
    }
    data.push(m);
}
export default data;