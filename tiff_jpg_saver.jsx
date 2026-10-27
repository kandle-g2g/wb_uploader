// tiff_jpg_saver v1.1
// G2G Code
//
// Утилита-конвертер: берёт папку с PSD (в т.ч. CMYK), для каждого файла
// сохраняет рядом TIFF (со всеми слоями сведёнными, в исходном цветовом
// режиме) и JPG (переведённый в RGB). Папка выбирается вручную каждый раз —
// скрипт не привязан к конкретному проекту, поэтому можно использовать его
// и для подготовки исходников перед раскладкой их в source/white и
// source/dark, и для любых других PSD.

function convertPsdToTiffAndJpg() {
    var inputFolder = Folder.selectDialog("Выберите папку с PSD файлами для обработки");
    if (!inputFolder) {
        alert("Операция отменена.");
        return;
    }

    var fileList = inputFolder.getFiles(/\.psd$/i);
    if (fileList.length === 0) {
        alert("В выбранной папке не найдено PSD файлов.");
        return;
    }

    // Выходная папка = та же, что и входная (никаких подпапок)
    var outputFolder = inputFolder;

    var tiffOptions = new TiffSaveOptions();
    tiffOptions.compression = TIFFEncoding.TIFFLZW;
    tiffOptions.embedColorProfile = true;
    tiffOptions.layers = false; // сводим слои

    var jpgOptions = new JPEGSaveOptions();
    jpgOptions.quality = 12;
    jpgOptions.embedColorProfile = true;
    jpgOptions.formatOptions = FormatOptions.STANDARDBASELINE;
    jpgOptions.matte = MatteType.NONE;
    jpgOptions.optimize = true;

    var done = 0;
    for (var i = 0; i < fileList.length; i++) {
        var file = fileList[i];
        var baseName = file.name.replace(/\.[^.]+$/, "");
        var doc = null;
        try {
            doc = app.open(file);

            // ---- TIFF в исходном цветовом режиме ----
            var tiffFile = new File(outputFolder + "/" + baseName + ".tif");
            doc.saveAs(tiffFile, tiffOptions, true, Extension.LOWERCASE);

            // ---- переводим в RGB перед JPG (JPEG не поддерживает CMYK) ----
            if (doc.mode === DocumentMode.CMYK) {
                doc.changeMode(ChangeMode.RGB);
            }

            // ---- JPG ----
            var jpgFile = new File(outputFolder + "/" + baseName + ".jpg");
            doc.saveAs(jpgFile, jpgOptions, true, Extension.LOWERCASE);

            done++;
        } catch (e) {
            alert("Ошибка при обработке \"" + file.name + "\": " + e.message + "\nФайл пропущен, продолжаю.");
        } finally {
            if (doc) doc.close(SaveOptions.DONOTSAVECHANGES);
        }
    }

    alert("Обработка завершена!\nОбработано файлов: " + done + " из " + fileList.length +
        "\nTIFF и JPG сохранены в:\n" + outputFolder.fsName);
}

if (BridgeTalk.appName === "photoshop") {
    convertPsdToTiffAndJpg();
} else {
    alert("Этот скрипт должен быть запущен из Adobe Photoshop.");
}