importScripts('./lib.min.js');
self.onmessage = function (e) {
    var options = e.data.options;
    var config = e.data.config;
    self.__exportExcel(options,config).then(function(buffer){
        self.postMessage(buffer);
    }).catch(function (error) {
        console.error(error);
        self.postMessage(-1);
    });
};