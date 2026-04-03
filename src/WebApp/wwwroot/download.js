window.downloadFile = function (fileName, contentType, base64Data) {
    // Convert base64 to binary using Blob (avoids data URL size limits and corruption)
    var byteCharacters = atob(base64Data);
    var byteNumbers = new Uint8Array(byteCharacters.length);
    for (var i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    var blob = new Blob([byteNumbers], { type: contentType });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.download = fileName;
    link.href = url;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
};
