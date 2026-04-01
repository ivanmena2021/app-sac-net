window.downloadFile = function (fileName, contentType, base64Data) {
    var link = document.createElement('a');
    link.download = fileName;
    link.href = 'data:' + contentType + ';base64,' + base64Data;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};
