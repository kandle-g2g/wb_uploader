// =====================================================================================
// MOCKUP FILLER — автоматическое заполнение мокапов постеров исходниками
// =====================================================================================
// Ожидаемая структура проекта (корневая папка):
//   Постеры.xlsx
//   mockups/     — cover.psd, 1.psd, 2.psd, 3.psd, 4.psd, 5-3.psd, 5-4.psd, 5-5.psd, 7.psd, 8.psd
//   output/      — сюда складываются результаты: output/<Исполнитель>/<исходник>/*.jpg
//   scripts/     — сам этот скрипт лежит здесь (используется для авто-поиска корня)
//   source/      — white/, dark/, "Other Albums"/
//
// Корневая папка определяется автоматически по расположению скрипта
// (scripts/.. = корень). Если авто-поиск не сработал (скрипт запущен из другого
// места или структура не совпадает) — показывается диалог ручного выбора папки.
//
// Как устроена обработка одного исходника (файла из source/white или source/dark):
//   1/2/3/7/8        — картинка вставляется в единственный смарт-объект мокапа;
//   1                — дополнительно подставляются текстовые слои Artist и Album Name;
//   2                — после подстановки картинки скрипт останавливается и включает
//                      Free Transform на слое лупы ("изображение"): подвиньте слой
//                      вручную и нажмите Enter (Esc — оставить как в шаблоне).
//                      Сдвиг запоминается для альбома и автоматически применяется
//                      к остальным расцветкам этого же альбома (white/dark и т.д.);
//   4                — вставляется сам исходник + подобранный "партнёр" того же
//                      альбома из противоположной папки (white<->dark); если
//                      партнёр не найден — мокап 4 для этого исходника
//                      пропускается целиком (как и мокапы 5-x при нехватке
//                      других альбомов), а не сохраняется наполовину пустым;
//   5-3/5-4/5-5      — подставляются другие альбомы исполнителя из "Other Albums"
//                      (текущий исходник туда не входит, кроме случая, когда других
//                      альбомов ровно 2 — тогда текущий добавляется третьим).
// Перед началом обработки показывается диалог со списком распознанных названий
// альбомов (можно поправить вручную, не трогая остальной текст строки).
// В этом же диалоге сверху есть чекбокс «Только JSON». Если его включить —
// скрипт спросит исполнителя, сохранит postery_album_export.json (для
// excel_filler.py) и завершится, ни разу не открыв ни один мокап-PSD. Так
// можно быстро пересобрать JSON (например, поправить название альбома),
// не дожидаясь и не прерывая полную сборку мокапов.
// =====================================================================================

#target photoshop
app.displayDialogs = DialogModes.NO;

// Список некритичных проблем, накопленных за время обработки (см. logIssue).
// Вместо alert() на каждую проблему по отдельности (что останавливает
// безнадзорный прогон скрипта до нажатия ОК) все такие сообщения
// собираются сюда и показываются одним alert'ом в самом конце (см. main()).
var ISSUES = [];
function logIssue(msg) {
    ISSUES.push(msg);
}

// Запомненные положения слоя лупы в мокапе 2: ключ — название альбома (из имени
// файла, без учёта регистра), значение — {dx, dy} сдвиг в px относительно положения
// в шаблоне. Живёт только в течение одного запуска скрипта.
var LOUPE_OFFSETS = {};

// ================================ КОНФИГ ============================================
var CONFIG = {
    // Имена файлов мокапов — ищутся напрямую внутри папки mockups/
    mockupFiles: {
        "1": "1.psd",
        "2": "2.psd",
        "3": "3.psd",
        "4": "4.psd",
        "6": "6.psd",
        "7": "7.psd",
        "5-3": "5-3.psd",
        "5-4": "5-4.psd",
        "5-5": "5-5.psd"
    },

    // Имена смарт-объектов внутри мокапов
    soNameSimple: "Smart Object",        // мокапы 1, 2, 3, 7
    soNameWhite: "Smart Object white",   // мокап 4
    soNameDark: "Smart Object dark",     // мокап 4
    soPrefix5x: "Smart Object ",         // мокапы 5-3/5-4/5-5: "Smart Object 1", "Smart Object 2"...

    // Названия папок (все — прямые подпапки корневой)
    mockupsFolderName: "mockups",
    sourceFolderName: "source",
    whiteFolderName: "white",
    darkFolderName: "dark",
    otherAlbumsFolderName: "Other Albums", // подпапка внутри source
    scriptsFolderName: "scripts", // сюда же кладём файл-передатчик для excel_filler.py

    // ДОПУЩЕНИЕ: реальных имён текстовых слоёв в мокапе 1 я не знаю (это было
    // "текущим функционалом" вашего старого скрипта, которого у меня нет).
    // Проверьте, как называются слои с именем исполнителя и альбома, и поправьте здесь.
    artistLayerName: "Artist",
    albumLayerName: "Album Name",

    // Мокапы 5-3/5-4/5-5: отдельный смарт-объект с именем исполнителя (с автоподбором
    // кегля, как у "Album Name" на мокапе 1) - другое имя слоя и свой максимальный кегль.
    artistNameLayerName: "Artist Name",
    artistNameMaxFontSize: 115, // максимальный кегль для "Artist Name" на мокапах 5-x, pt

    // Мокап 2: слой лупы (копия смарт-объекта), положение которого настраивается
    // вручную для каждого альбома. Имя ищется без учёта регистра.
    loupeLayerName: "изображение",

    // Как вписывать исходник в холст смарт-объекта:
    // "cover"   — масштабировать так, чтобы исходник полностью залил холст
    //             (возможна обрезка по краям, без искажения пропорций);
    // "contain" — вписать целиком без обрезки (могут остаться пустые поля).
    fitMode: "cover",
    // Небольшой запас при cover-масштабировании (в долях, 0.01 = +1%), чтобы из-за
    // округления границ слоя в пикселях не оставалось тонкой щели по краю холста.
    // Это применяется ВНУТРИ плоского (неискажённого) документа смарт-объекта, так
    // что на итоговую перспективу мокапа никак не влияет и безопасно в любую сторону.
    coverOverscan: 0.005,

    // Расширения исходников
    extensions: ["jpg", "jpeg"],

    // Автоподбор кегля для текста внутри смарт-объекта "Album Name"
    // (взято из вашего рабочего скрипта fit-album-name)
    maxFontSize: 135,   // максимальный кегль, pt
    minFontSize: 8,      // минимальный кегль, pt (страховка)
    searchSteps: 25,     // шагов бинарного поиска кегля
    leadingRatio: 0.9,   // межстрочный интервал = размер шрифта × это число

    // Результат сохраняется в: <корень>/output/<исполнитель>/<исходник без расширения>/<id мокапа>.jpg
    outputFolderName: "output",
    jpegQuality: 12,

    // Имя файла, через который подтверждённые названия альбомов передаются
    // в excel_filler.py. Кладётся в папку scripts/ (рядом с самими скриптами,
    // на уровень ниже Постеры.xlsx), перезаписывается при каждом запуске и
    // удаляется python-скриптом после использования — мусора не остаётся.
    albumExportFileName: "postery_album_export.json",

    // Общий конфиг проекта (лежит в scripts/). Сейчас mockuper использует из него только artist,
    // если он указан. Остальные настройки (например, токен WB) Photoshop
    // вообще не читает. Это позволяет будущему Posters.app задавать исполнителя
    // без второго окна Photoshop, сохраняя ручной ввод как запасной вариант.
    projectConfigFileName: "config.json"
};

// ================================ УТИЛИТЫ: ФАЙЛЫ/ПАПКИ ==============================

// Ищет файл СРАЗУ внутри указанной папки (без рекурсии — расположение всех
// мокапов внутри mockups/ фиксировано и известно заранее).
function findFileInFolder(folder, fileName) {
    if (!folder) return null;
    var items = folder.getFiles();
    for (var i = 0; i < items.length; i++) {
        if (items[i] instanceof File &&
            normalizeForCompare(decodeFileName(items[i].name)) === normalizeForCompare(fileName)) {
            return items[i];
        }
    }
    return null;
}

// Сравнение имён папок/файлов не только без учёта регистра, но и без учёта
// лишних пробелов и неразрывного пробела (\u00A0), а также URL-кодирования
// (%20 и т.п.) — на диске реальные имена иногда оказываются закодированы так.
function normalizeForCompare(str) {
    return str.replace(/[\s\u00A0]+/g, " ").replace(/^\s+|\s+$/g, "").toLowerCase();
}

function getDirectSubfolder(parent, name) {
    if (!parent) return null;
    var items = parent.getFiles();
    for (var i = 0; i < items.length; i++) {
        if (items[i] instanceof Folder &&
            normalizeForCompare(decodeFileName(items[i].name)) === normalizeForCompare(name)) {
            return items[i];
        }
    }
    return null;
}

function hasAllowedExtension(file) {
    var dot = file.name.lastIndexOf(".");
    if (dot === -1) return false;
    var ext = file.name.substring(dot + 1).toLowerCase();
    for (var i = 0; i < CONFIG.extensions.length; i++) {
        if (ext === CONFIG.extensions[i]) return true;
    }
    return false;
}

function listImageFiles(folder) {
    var result = [];
    if (!folder) return result;
    var items = folder.getFiles();
    for (var i = 0; i < items.length; i++) {
        if (items[i] instanceof File && hasAllowedExtension(items[i])) {
            result.push(items[i]);
        }
    }
    result.sort(function (a, b) {
        return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
    });
    return result;
}

function stripExtension(name) {
    var dot = name.lastIndexOf(".");
    return dot > -1 ? name.substring(0, dot) : name;
}

// ExtendScript иногда возвращает File.name в URL-кодированном виде
// (пробелы как %20 и т.п.), особенно если в пути есть пробелы или не-ASCII
// символы. Декодируем перед любым текстовым разбором имени.
function decodeFileName(name) {
    try {
        return decodeURI(name);
    } catch (e) {
        return name;
    }
}

function ensureFolder(folder) {
    if (!folder.exists) folder.create();
    return folder;
}

// ================================ ЭКСПОРТ НАЗВАНИЙ ДЛЯ excel_filler.py ===============
// Пишем простой JSON вручную (не полагаемся на глобальный JSON — в части версий
// ExtendScript его нет), чтобы python-скрипт мог подхватить те же названия
// альбомов, что вы подтвердили в диалоге, без повторного ввода вручную.
function jsonEscape(str) {
    return String(str)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\r\n|\r|\n/g, "\\n");
}

// albumMap: { "имя файла без расширения": "название альбома", ... }
// scriptsFolder: папка scripts/ внутри проекта (где лежат сами скрипты) -
// файл кладём туда, рядом с excel_filler.py, чтобы не мусорить в корне.
function writeAlbumExportFile(scriptsFolder, artistName, albumMap) {
    var parts = [];
    for (var key in albumMap) {
        if (!albumMap.hasOwnProperty(key)) continue;
        parts.push('"' + jsonEscape(key) + '":"' + jsonEscape(albumMap[key]) + '"');
    }
    var json = '{"artist":"' + jsonEscape(artistName) + '","albums":{' + parts.join(",") + '}}';

    var exportFile = new File(scriptsFolder.fsName + "/" + CONFIG.albumExportFileName);
    exportFile.encoding = "UTF-8";
    exportFile.open("w");
    exportFile.write(json);
    exportFile.close();
}

// ================================ РАЗБОР ИМЕНИ ФАЙЛА ================================
// Шаблон: "<альбом, слово+> <цвет — ровно одно слово>"
// Имени исполнителя в имени файла нет, поэтому просто отрезаем последнее слово —
// это цвет, всё, что до него — название альбома.

function deriveAlbumAndColor(baseName) {
    var trimmed = baseName.replace(/\s+$/, "");
    var lastSpace = trimmed.lastIndexOf(" ");
    if (lastSpace === -1) return null; // не нашли отдельного слова цвета

    return {
        album: trimmed.substring(0, lastSpace),
        color: trimmed.substring(lastSpace + 1)
    };
}

// ================================ ДИАЛОГИ ============================================

// ================================ КОНФИГ ПРОЕКТА ===================================
// Читаем только поле artist из scripts/config.json.
// Токен WB и любые другие секреты здесь не используются и не выводятся.
function readProjectArtist(rootFolder) {
    try {
        var configFile = new File(rootFolder.fsName + "/" + CONFIG.scriptsFolderName + "/" + CONFIG.projectConfigFileName);
        if (!configFile.exists) return null;

        configFile.encoding = "UTF-8";
        if (!configFile.open("r")) return null;
        var text = configFile.read();
        configFile.close();

        // Поддерживаем оба варианта:
        // {"project":{"artist":"..."}}
        // {"artist":"..."}
        var match = text.match(/"project"\s*:\s*\{[\s\S]*?"artist"\s*:\s*"((?:\\.|[^"\\])*)"/);
        if (!match) {
            match = text.match(/"artist"\s*:\s*"((?:\\.|[^"\\])*)"/);
        }
        if (!match || !match[1]) return null;

        return match[1]
            .replace(/\\"/g, '"')
            .replace(/\\\\/g, "\\")
            .replace(/\\n/g, "\n")
            .replace(/\\r/g, "\r")
            .replace(/\\t/g, "\t")
            .replace(/^\s+|\s+$/g, "");
    } catch (e) {
        return null;
    }
}

// Берём исполнителя из config.json, а если его там нет — сохраняем старый
// интерактивный сценарий с окном Photoshop. Поэтому скрипт остаётся совместимым
// с ручным запуском даже без нового приложения.
function getArtistName(rootFolder) {
    var fromConfig = readProjectArtist(rootFolder);
    if (fromConfig) return fromConfig;

    return promptText("Исполнитель", "Введите Имя исполнителя:", "");
}

function promptText(title, message, defaultValue) {
    var w = new Window("dialog", title);
    w.orientation = "column";
    w.alignChildren = "fill";
    w.add("statictext", undefined, message);
    var input = w.add("edittext", undefined, defaultValue || "");
    input.characters = 40;
    input.active = true;

    var buttons = w.add("group");
    buttons.alignment = "right";
    var cancelBtn = buttons.add("button", undefined, "Cancel", { name: "cancel" });
    var okBtn = buttons.add("button", undefined, "OK", { name: "ok" });

    var result = null;
    okBtn.onClick = function () { result = input.text; w.close(); };
    cancelBtn.onClick = function () { result = null; w.close(); };
    w.show();
    return result;
}

// ---------- Строка предпросмотра для одного исходника ----------
// Формат: "Распознанное название альбома" - Исходное имя файла.jpg
// Если имя файла не разбирается на альбом+цвет (parsed === null), в кавычках
// показывается имя файла целиком без расширения — как понятный запасной
// вариант, который можно тут же поправить руками.
function buildAlbumPreviewLine(sourceFile) {
    var baseName = stripExtension(decodeFileName(sourceFile.name));
    var parsed = deriveAlbumAndColor(baseName);
    var albumGuess = parsed ? parsed.album : baseName;
    return "\"" + albumGuess + "\" - " + decodeFileName(sourceFile.name);
}

// ---------- Диалог проверки/правки названий альбомов перед запуском ----------
// sources — массив {file, side}, в том же порядке, что и в главном цикле.
// Возвращает отредактированный пользователем текст (или null при отмене).
function promptAlbumReview(sources) {
    var lines = [];
    for (var i = 0; i < sources.length; i++) {
        lines.push(buildAlbumPreviewLine(sources[i].file));
    }
    var initialText = lines.join("\n");

    var w = new Window("dialog", "Проверка названий альбомов");
    w.orientation = "column";
    w.alignChildren = "fill";

    var modePanel = w.add("panel", undefined, "Режим запуска");
    modePanel.orientation = "column";
    modePanel.alignChildren = "left";
    modePanel.margins = [10, 15, 10, 10];
    var jsonOnlyCheck = modePanel.add("checkbox", undefined,
        "Только JSON — сохранить альбомы/исполнителя для excel_filler.py и завершить, без сборки мокапов");
    jsonOnlyCheck.helpTip = "Мокап-PSD не открываются вообще — Photoshop не занят и не нужно " +
        "принудительно закрывать его между запусками.";

    w.add("statictext", undefined,
        "Так названия альбомов будут подставлены в мокап 1 (слой \"" + CONFIG.albumLayerName + "\").");
    w.add("statictext", undefined,
        "Формат строки: \"название альбома\" - имя исходного файла");
    w.add("statictext", undefined,
        "Правьте только текст В КАВЫЧКАХ. Не добавляйте, не удаляйте и не меняйте местами строки.");
    w.add("statictext", undefined,
        "Вид кавычек (\" или «ёлочки») не важен — редактор может подменить их сам, это ОК.");

    var input = w.add("edittext", undefined, initialText, { multiline: true, scrolling: true });
    input.preferredSize = [640, 420];

    var buttons = w.add("group");
    buttons.alignment = "right";
    var cancelBtn = buttons.add("button", undefined, "Отмена", { name: "cancel" });
    var okBtn = buttons.add("button", undefined, "OK", { name: "ok" });

    var result = null;
    okBtn.onClick = function () {
        result = { text: input.text, jsonOnly: jsonOnlyCheck.value };
        w.close();
    };
    cancelBtn.onClick = function () { result = null; w.close(); };
    w.show();
    return result;
}

// ---------- Разбор отредактированного текста обратно в список альбомов ----------
// Из каждой строки берётся значение в первых кавычках (по порядку строк).
// Количество строк должно точно совпадать с количеством исходников — иначе
// сопоставление строка-исходник может съехать, поэтому в этом случае
// выполнение останавливается с понятным сообщением.
//
// ВАЖНО про кавычки: при вводе текста в поле диалога macOS/Windows часто сами
// подменяют прямую кавычку " на типографскую («ёлочки» « », "лапки" “ ” и т.п.).
// Поэтому открывающей и закрывающей кавычкой в строке может оказаться любой
// символ из набора ниже — причём открывающая и закрывающая необязательно
// одного и того же вида. Разбор строки на это не завязан.
var QUOTE_CHARS = "\"\u00AB\u00BB\u201C\u201D\u201E\u201F"; // " « » “ ” „ ‟
var QUOTE_CLASS = "[" + QUOTE_CHARS + "]";
var albumLineRe = new RegExp("^\\s*" + QUOTE_CLASS + "([^" + QUOTE_CHARS + "]*)" + QUOTE_CLASS);

function parseAlbumOverrides(text, expectedCount) {
    var lines = text.split(/\r\n|\r|\n/);
    // убираем пустые строки в конце текста (их часто добавляет сам редактор поля)
    while (lines.length && /^\s*$/.test(lines[lines.length - 1])) lines.pop();

    if (lines.length !== expectedCount) {
        alert("Количество строк в диалоге (" + lines.length + ") не совпадает с количеством " +
            "исходников (" + expectedCount + ").\nПравьте только текст в кавычках, не добавляйте " +
            "и не удаляйте строки.\nОстановка.");
        return null;
    }

    var albums = [];
    for (var i = 0; i < lines.length; i++) {
        var m = lines[i].match(albumLineRe);
        if (!m || normalizeForCompare(m[1]) === "") {
            alert("Не удалось разобрать строку " + (i + 1) + ": \"" + lines[i] + "\"\n" +
                "Ожидался формат: \"название альбома\" - имя файла (кавычки любые)\nОстановка.");
            return null;
        }
        albums.push(m[1]);
    }
    return albums;
}

// ================================ PHOTOSHOP: СЛОИ ====================================

function findLayerRecursive(layers, name) {
    for (var i = 0; i < layers.length; i++) {
        var layer = layers[i];
        if (layer.name === name) return layer;
        if (layer.typename === "LayerSet") {
            var found = findLayerRecursive(layer.layers, name);
            if (found) return found;
        }
    }
    return null;
}

function withPixelRulerUnits(fn) {
    var original = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    try {
        return fn();
    } finally {
        app.preferences.rulerUnits = original;
    }
}

// ---------- Войти в редактирование смарт-объекта ----------
// На части версий Photoshop ArtLayer.editContents() недоступен в скриптовом DOM —
// используем прямой вызов того же меню-действия через Action Manager.
function editSmartObjectContents(doc, layer) {
    doc.activeLayer = layer;
    try {
        layer.editContents();
        return;
    } catch (e) {
        // fallback ниже
    }
    var idPlacedLayerEditContents = stringIDToTypeID("placedLayerEditContents");
    executeAction(idPlacedLayerEditContents, undefined, DialogModes.NO);
}

// ---------- Подогнать слой под размер targetWidth x targetHeight (px), с центрированием ----------
function fitLayerToCanvas(layer, targetWidth, targetHeight, fitMode) {
    var b = layer.bounds;
    var w = b[2].value - b[0].value;
    var h = b[3].value - b[1].value;
    if (w <= 0 || h <= 0) return;

    var scaleX = targetWidth / w;
    var scaleY = targetHeight / h;
    var scale = (fitMode === "contain") ? Math.min(scaleX, scaleY) : Math.max(scaleX, scaleY);
    if (fitMode !== "contain") {
        scale *= (1 + CONFIG.coverOverscan);
    }

    layer.resize(scale * 100, scale * 100, AnchorPosition.MIDDLECENTER);

    var ab = layer.bounds;
    var curCx = (ab[0].value + ab[2].value) / 2;
    var curCy = (ab[1].value + ab[3].value) / 2;
    layer.translate(targetWidth / 2 - curCx, targetHeight / 2 - curCy);
}

// ---------- Скопировать пиксели файла в целевой документ как новый слой ----------
// Через duplicate() между документами (а не paste), чтобы результат не зависел
// от несовпадения PPI/разрешения между исходником и документом смарт-объекта —
// пиксельные размеры переносятся как есть.
function placeImageAsLayer(targetDoc, file) {
    var srcDoc = app.open(file);
    try {
        if (srcDoc.layers.length > 1) srcDoc.flatten();
        var dup = srcDoc.layers[0].duplicate(targetDoc, ElementPlacement.PLACEATBEGINNING);
        return dup;
    } finally {
        srcDoc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

// ---------- Удалить все слои документа, кроме указанного ----------
function clearAllLayersExcept(doc, keepLayer) {
    for (var i = doc.layers.length - 1; i >= 0; i--) {
        var l = doc.layers[i];
        if (keepLayer && l.id === keepLayer.id) continue;
        try { l.remove(); } catch (e) {}
    }
}

// ---------- Вставить картинку в смарт-объект, редактируя его содержимое напрямую ----------
// Это аналог ручной вставки (двойной клик → вставить/подогнать → Save). Внешняя
// трансформация/перспектива плейсхолдера в родительском документе не трогается —
// поэтому окантовки/вылезаний по краям больше не возникает, даже на искажённых
// мокапах.
function fillSmartObjectWithImage(mainDoc, layer, file, fitMode) {
    withPixelRulerUnits(function () {
        editSmartObjectContents(mainDoc, layer);
        var soDoc = app.activeDocument;
        try {
            var targetWidth = soDoc.width.as("px");
            var targetHeight = soDoc.height.as("px");

            var imgLayer = placeImageAsLayer(soDoc, file);
            fitLayerToCanvas(imgLayer, targetWidth, targetHeight, fitMode);
            clearAllLayersExcept(soDoc, imgLayer);

            soDoc.close(SaveOptions.SAVECHANGES);
        } catch (e) {
            try { soDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e2) {}
            throw e;
        }
        app.activeDocument = mainDoc;
    });
}

// ---------- Первый текстовый слой (в т.ч. в группах) ----------
function findFirstTextLayer(layers) {
    for (var i = 0; i < layers.length; i++) {
        var l = layers[i];
        if (l.kind === LayerKind.TEXT) return l;
        if (l.typename === "LayerSet") {
            var found = findFirstTextLayer(l.layers);
            if (found) return found;
        }
    }
    return null;
}

// ---------- Снять стиль (шрифт/цвет/трекинг) с первого текстового слоя ----------
function captureTextStyle(layers) {
    for (var i = 0; i < layers.length; i++) {
        var l = layers[i];
        if (l.kind === LayerKind.TEXT) {
            var ti = l.textItem;
            var style = {};
            try { style.font = ti.font; } catch (e) {}
            try { style.color = ti.color; } catch (e) {}
            try { style.capitalization = ti.capitalization; } catch (e) {}
            try { style.tracking = ti.tracking; } catch (e) {}
            return style;
        }
        if (l.typename === "LayerSet") {
            var found = captureTextStyle(l.layers);
            if (found) return found;
        }
    }
    return null;
}

// ---------- Подгонка текстового слоя под область width x height (px) ----------
function fitTextLayerToBox(layer, targetWidth, targetHeight, maxFontSize) {
    maxFontSize = (maxFontSize !== undefined) ? maxFontSize : CONFIG.maxFontSize;
    var ti = layer.textItem;

    ti.kind = TextType.PARAGRAPHTEXT;
    ti.width = new UnitValue(targetWidth, "px");
    ti.height = new UnitValue(targetHeight * 10, "px"); // с запасом, не влияет на расчёт
    ti.position = [0, 0];
    ti.justification = Justification.LEFT;
    ti.useAutoLeading = false; // авто-интерлиньяж слишком широкий на крупном кегле

    function fits(size) {
        ti.size = size;
        ti.leading = new UnitValue(size * CONFIG.leadingRatio, "pt");
        var b = layer.bounds;
        var w = b[2].value - b[0].value;
        var h = b[3].value - b[1].value;
        return (w <= targetWidth + 0.5 && h <= targetHeight + 0.5);
    }

    var best = CONFIG.minFontSize;

    if (fits(maxFontSize)) {
        best = maxFontSize;
    } else {
        var low = CONFIG.minFontSize, high = maxFontSize;
        for (var i = 0; i < CONFIG.searchSteps; i++) {
            var mid = (low + high) / 2;
            if (fits(mid)) { best = mid; low = mid; } else { high = mid; }
        }
    }

    var safety = 0;
    while (!fits(best) && best > CONFIG.minFontSize && safety < 300) {
        best -= 0.1;
        safety++;
    }

    ti.size = best;

    // Центрирование по вертикали
    var b = layer.bounds;
    var renderedHeight = b[3].value - b[1].value;
    var currentTop = b[1].value;
    var desiredTop = (targetHeight - renderedHeight) / 2;
    var shiftY = desiredTop - currentTop;
    layer.translate(0, shiftY);

    return best;
}

// ---------- Заменить содержимое слоя "Artist" (без подгонки размера) ----------
// Слой может быть обычным текстовым слоем ИЛИ смарт-объектом с текстом внутри —
// обрабатываются оба случая.
function replaceArtistLayerContent(mainDoc, layerName, newText) {
    var layer = findLayerRecursive(mainDoc.layers, layerName);
    if (!layer) {
        logIssue("Не найден слой \"" + layerName + "\" — имя исполнителя не подставлено.");
        return;
    }

    if (layer.kind === LayerKind.TEXT) {
        layer.textItem.contents = newText;
        return;
    }

    if (layer.kind === LayerKind.SMARTOBJECT) {
        editSmartObjectContents(mainDoc, layer);
        var soDoc = app.activeDocument;
        try {
            var textLayer = findFirstTextLayer(soDoc.layers);
            if (!textLayer) {
                logIssue("Внутри смарт-объекта \"" + layerName + "\" не найден текстовый слой.");
            } else {
                textLayer.textItem.contents = newText;
            }
            soDoc.close(SaveOptions.SAVECHANGES);
        } catch (e) {
            try { soDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e2) {}
            throw e;
        }
        app.activeDocument = mainDoc;
        return;
    }

    logIssue("Слой \"" + layerName + "\" найден, но это не текстовый слой и не смарт-объект — пропущено.");
}

// ---------- Заменить содержимое смарт-объекта с автоподбором кегля ----------
// Общая логика для "Album Name" (мокап 1) и "Artist Name" (мокапы 5-x) - разница
// только в имени слоя и максимальном кегле.
//
// ВАЖНО про порядок: шрифт (ti.font) применяется ДО того, как задаётся текст
// (ti.contents). Раньше было наоборот, и из-за этого Photoshop успевал "напечатать"
// текст шрифтом по умолчанию, а для символов вроде "ё"/"й" — которых нет в этом
// дефолтном шрифте — переключался на системный fallback ещё ДО того, как скрипт
// применял нужный шрифт; сама смена ti.font после этого такие символы уже не
// перекрашивала. При ручном наборе в Photoshop нужный шрифт выбран заранее,
// поэтому там всё было верно с самого начала.
function replaceTextSmartObjectWithAutoFit(mainDoc, layerName, newText, maxFontSize) {
    var soLayer = findLayerRecursive(mainDoc.layers, layerName);
    if (!soLayer) {
        logIssue("Не найден слой \"" + layerName + "\" — текст не подставлен.");
        return;
    }
    if (soLayer.kind !== LayerKind.SMARTOBJECT) {
        logIssue("Слой \"" + layerName + "\" найден, но это не смарт-объект — текст не подставлен.");
        return;
    }

    editSmartObjectContents(mainDoc, soLayer);
    var soDoc = app.activeDocument;
    try {
        var targetWidth = soDoc.width.as("px");
        var targetHeight = soDoc.height.as("px");

        var oldStyle = captureTextStyle(soDoc.layers);

        var textLayer = soDoc.artLayers.add();
        textLayer.kind = LayerKind.TEXT;
        var ti = textLayer.textItem;

        // Стиль (в первую очередь шрифт) — ДО контента, см. комментарий выше.
        if (oldStyle) {
            if (oldStyle.font) { try { ti.font = oldStyle.font; } catch (e) {} }
            if (oldStyle.color) { try { ti.color = oldStyle.color; } catch (e) {} }
            if (oldStyle.capitalization !== undefined) { try { ti.capitalization = oldStyle.capitalization; } catch (e) {} }
            if (oldStyle.tracking !== undefined) { try { ti.tracking = oldStyle.tracking; } catch (e) {} }
        }

        ti.contents = newText;

        clearAllLayersExcept(soDoc, textLayer);
        fitTextLayerToBox(textLayer, targetWidth, targetHeight, maxFontSize);

        soDoc.close(SaveOptions.SAVECHANGES);
    } catch (e) {
        try { soDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (e2) {}
        throw e;
    }
    app.activeDocument = mainDoc;
}

// ---------- Заменить содержимое смарт-объекта "Album Name" (мокап 1) ----------
function replaceAlbumNameSmartObject(mainDoc, layerName, newText) {
    replaceTextSmartObjectWithAutoFit(mainDoc, layerName, newText, CONFIG.maxFontSize);
}

function exportJPEG(doc, file, quality) {
    var opts = new JPEGSaveOptions();
    opts.quality = quality;
    doc.flatten();
    doc.saveAs(file, opts, true, Extension.LOWERCASE);
}

// ================================ ОБРАБОТКА ОДНОГО МОКАПА ============================

function fillSimpleMockup(mockupFile, soName, sourceFile, outFile) {
    var doc = app.open(mockupFile);
    try {
        var so = findLayerRecursive(doc.layers, soName);
        if (!so) {
            logIssue("Не найден смарт-объект \"" + soName + "\" в файле " + decodeFileName(mockupFile.name));
        } else {
            fillSmartObjectWithImage(doc, so, sourceFile, CONFIG.fitMode);
        }
        exportJPEG(doc, outFile, CONFIG.jpegQuality);
    } finally {
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

// ---------- Поиск слоя по имени без учёта регистра ----------
function findLayerByNameCI(layers, name) {
    var target = normalizeForCompare(name);
    for (var i = 0; i < layers.length; i++) {
        var layer = layers[i];
        if (normalizeForCompare(layer.name) === target) return layer;
        if (layer.typename === "LayerSet") {
            var found = findLayerByNameCI(layer.layers, name);
            if (found) return found;
        }
    }
    return null;
}

// ---------- Free Transform в интерактивном режиме ----------
// Photoshop показывает рамку трансформации на активном слое и ждёт, пока
// пользователь подвинет слой и нажмёт Enter; скрипт при этом стоит на месте.
// Возвращает true при Enter, false при Esc (отмена пользователем).
function interactiveFreeTransform() {
    var savedDialogs = app.displayDialogs;
    app.displayDialogs = DialogModes.ALL; // глобально стоит NO — на время вызова включаем
    try {
        var desc = new ActionDescriptor();
        desc.putEnumerated(charIDToTypeID("FTcs"), charIDToTypeID("QCSt"), charIDToTypeID("Qcsa"));
        var offset = new ActionDescriptor();
        offset.putUnitDouble(charIDToTypeID("Hrzn"), charIDToTypeID("#Pxl"), 0);
        offset.putUnitDouble(charIDToTypeID("Vrtc"), charIDToTypeID("#Pxl"), 0);
        desc.putObject(charIDToTypeID("Ofst"), charIDToTypeID("Ofst"), offset);
        executeAction(charIDToTypeID("Trnf"), desc, DialogModes.ALL);
        return true;
    } catch (e) {
        if (e.number === 8007) return false; // "User cancelled" — нажали Esc
        throw e;
    } finally {
        app.displayDialogs = savedDialogs;
    }
}

// ---------- Положение слоя лупы в мокапе 2 ----------
// Первый раз для альбома: ручная подгонка (Free Transform, Enter), сдвиг запоминается.
// Дальше для этого же альбома (другая расцветка): тот же сдвиг применяется автоматически.
function positionLoupeLayer(doc, albumKey) {
    var layer = findLayerByNameCI(doc.layers, CONFIG.loupeLayerName);
    if (!layer) {
        logIssue("Не найден слой \"" + CONFIG.loupeLayerName + "\" в 2.psd — положение лупы не настроено");
        return;
    }

    withPixelRulerUnits(function () {
        var saved = LOUPE_OFFSETS[albumKey];
        if (saved) {
            if (saved.dx !== 0 || saved.dy !== 0) layer.translate(saved.dx, saved.dy);
            return;
        }

        app.activeDocument = doc;
        doc.activeLayer = layer;
        app.refresh(); // чтобы на экране уже была подставленная обложка

        var b0 = layer.bounds;
        var x0 = b0[0].value, y0 = b0[1].value;

        var confirmed;
        try {
            confirmed = interactiveFreeTransform();
        } catch (e) {
            logIssue("Не удалось запустить ручное перемещение слоя лупы (" + e.message + ") — оставлено положение из шаблона");
            return;
        }

        if (!confirmed) {
            // Esc — оставляем как в шаблоне и не спрашиваем снова для других расцветок этого альбома
            LOUPE_OFFSETS[albumKey] = { dx: 0, dy: 0 };
            return;
        }

        var b1 = layer.bounds;
        LOUPE_OFFSETS[albumKey] = { dx: b1[0].value - x0, dy: b1[1].value - y0 };
    });
}

// Мокап 2: картинка в смарт-объект + ручная подгонка слоя лупы (см. positionLoupeLayer).
function fillMockup2(mockupFile, soName, sourceFile, outFile, albumKey) {
    var doc = app.open(mockupFile);
    try {
        var so = findLayerRecursive(doc.layers, soName);
        if (!so) {
            logIssue("Не найден смарт-объект \"" + soName + "\" в файле " + decodeFileName(mockupFile.name));
        } else {
            fillSmartObjectWithImage(doc, so, sourceFile, CONFIG.fitMode);
        }

        positionLoupeLayer(doc, albumKey);

        exportJPEG(doc, outFile, CONFIG.jpegQuality);
    } finally {
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

// Мокап 1 отдельно: помимо картинки в нём есть слой "Artist" (текст или
// смарт-объект с текстом) и слой-смарт-объект "Album Name" с автоподбором кегля.
function fillMockup1(mockupFile, soImageName, sourceFile, outFile, artistName, albumName, artistLayerName, albumLayerName) {
    var doc = app.open(mockupFile);
    try {
        var so = findLayerRecursive(doc.layers, soImageName);
        if (!so) {
            logIssue("Не найден смарт-объект \"" + soImageName + "\" в файле " + decodeFileName(mockupFile.name));
        } else {
            fillSmartObjectWithImage(doc, so, sourceFile, CONFIG.fitMode);
        }

        replaceArtistLayerContent(doc, artistLayerName, artistName);
        replaceAlbumNameSmartObject(doc, albumLayerName, albumName);

        exportJPEG(doc, outFile, CONFIG.jpegQuality);
    } finally {
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

// ================================ ГЛАВНЫЙ СКРИПТ =====================================

// ---------- Определить корневую папку проекта ----------
// Скрипт лежит в <корень>/scripts/, поэтому по умолчанию корень — это папка на
// уровень выше той, где лежит сам скрипт. Если рядом не находится ожидаемых
// подпапок (mockups, source) — считаем, что скрипт запущен не оттуда, и просим
// выбрать корневую папку вручную.
function layoutLooksValid(folder) {
    return !!getDirectSubfolder(folder, CONFIG.mockupsFolderName) &&
        !!getDirectSubfolder(folder, CONFIG.sourceFolderName);
}

function resolveRootFolder() {
    try {
        var scriptFile = new File($.fileName);
        var candidateRoot = scriptFile.parent.parent; // scripts/.. 
        if (candidateRoot && candidateRoot.exists && layoutLooksValid(candidateRoot)) {
            return candidateRoot;
        }
    } catch (e) {}

    var picked = Folder.selectDialog(
        "Не удалось определить корневую папку проекта автоматически.\n" +
        "Выберите папку, где лежат Постеры.xlsx, mockups, output, scripts, source"
    );
    return picked;
}

function main() {
    var rootFolder = resolveRootFolder();
    if (!rootFolder) { alert("Отменено: корневая папка не определена."); return; }

    var sourceFolder = getDirectSubfolder(rootFolder, CONFIG.sourceFolderName);
    if (!sourceFolder) { alert("Не найдена папка \"" + CONFIG.sourceFolderName + "\" в " + rootFolder.fsName); return; }

    var whiteFolder = getDirectSubfolder(sourceFolder, CONFIG.whiteFolderName);
    var darkFolder = getDirectSubfolder(sourceFolder, CONFIG.darkFolderName);
    if (!whiteFolder || !darkFolder) {
        alert("В папке \"" + CONFIG.sourceFolderName + "\" не найдены подпапки \"" +
            CONFIG.whiteFolderName + "\" и/или \"" + CONFIG.darkFolderName + "\"");
        return;
    }

    var otherAlbumsFolder = getDirectSubfolder(sourceFolder, CONFIG.otherAlbumsFolderName);
    if (!otherAlbumsFolder) {
        alert("Папка \"" + CONFIG.otherAlbumsFolderName + "\" не найдена внутри \"" +
            CONFIG.sourceFolderName + "\".\nМокапы 5-х будут пропущены для ВСЕХ исходников.");
    }

    // собираем список исходников: white и dark вместе, с пометкой стороны
    var sources = [];
    var whiteFiles = listImageFiles(whiteFolder);
    var darkFiles = listImageFiles(darkFolder);
    for (var i = 0; i < whiteFiles.length; i++) sources.push({ file: whiteFiles[i], side: "white" });
    for (var j = 0; j < darkFiles.length; j++) sources.push({ file: darkFiles[j], side: "dark" });

    if (sources.length === 0) { alert("В папках white/dark не найдено исходников"); return; }

    // Диалог проверки названий альбомов — до начала любой обработки.
    // Показывает по строке на каждый исходник: "распознанное название" - имя файла.
    // Позволяет заранее поправить название альбома, если разбор имени файла дал
    // что-то не то — правится текст в кавычках, порядок строк = порядку sources.
    // Сверху в этом же диалоге — чекбокс «Только JSON» (режим работы).
    var albumReview = promptAlbumReview(sources);
    if (albumReview === null) { alert("Отменено на этапе проверки названий альбомов."); return; }

    var albumOverrides = parseAlbumOverrides(albumReview.text, sources.length);
    if (!albumOverrides) return; // сообщение об ошибке уже показано внутри

    for (var ai = 0; ai < sources.length; ai++) {
        sources[ai].albumOverride = albumOverrides[ai];
    }

    // Имя исполнителя берём из config.json, если его записало приложение.
    // При ручном запуске без artist в конфиге остаётся старое окно Photoshop.
    var artistName = getArtistName(rootFolder);
    if (!artistName) { alert("Имя исполнителя не указано — остановка"); return; }

    // Экспортируем подтверждённые названия альбомов для excel_filler.py —
    // ключ: имя файла-исходника без расширения (совпадает с тем, как его
    // видит python-скрипт), значение: название альбома из диалога проверки.
    // Это происходит в ОБОИХ режимах — в режиме «Только JSON» это единственный
    // результат работы скрипта.
    try {
        var scriptsFolder = ensureFolder(new Folder(rootFolder.fsName + "/" + CONFIG.scriptsFolderName));
        var albumExportMap = {};
        for (var mi = 0; mi < sources.length; mi++) {
            var stemKey = stripExtension(decodeFileName(sources[mi].file.name));
            albumExportMap[stemKey] = sources[mi].albumOverride;
        }
        writeAlbumExportFile(scriptsFolder, artistName, albumExportMap);
    } catch (eExport) {
        alert("Не удалось сохранить названия альбомов для Excel-скрипта: " + eExport.message +
            (albumReview.jsonOnly ? "" :
                "\nМокапы всё равно будут собраны, но в Excel названия придётся вписать вручную."));
        if (albumReview.jsonOnly) return; // в этом режиме сохранить JSON и было единственной целью
    }

    // Режим «Только JSON»: файл уже сохранён выше — останавливаемся здесь и не
    // трогаем ни один мокап-PSD, поэтому Photoshop не занят открытыми документами
    // и не требует принудительного закрытия между запусками.
    if (albumReview.jsonOnly) {
        alert("Готово. Сохранён только " + CONFIG.albumExportFileName + " (" + sources.length +
            " исходников, исполнитель \"" + artistName + "\"). Мокапы не собирались.");
        return;
    }

    // находим все мокапы (напрямую внутри mockups/, без рекурсии) — только теперь,
    // когда точно известно, что понадобится полная сборка
    var mockupsFolder = getDirectSubfolder(rootFolder, CONFIG.mockupsFolderName);
    if (!mockupsFolder) { alert("Не найдена папка \"" + CONFIG.mockupsFolderName + "\" в " + rootFolder.fsName); return; }

    var mockupPaths = {};
    for (var key in CONFIG.mockupFiles) {
        var found = findFileInFolder(mockupsFolder, CONFIG.mockupFiles[key]);
        if (!found) { alert("Не найден мокап " + CONFIG.mockupFiles[key] + " в папке \"" + CONFIG.mockupsFolderName + "\""); return; }
        mockupPaths[key] = found;
    }

    var outputRoot = ensureFolder(new Folder(rootFolder.fsName + "/" + CONFIG.outputFolderName + "/" + artistName));

    var processed = 0;

    // -------------------------------------------------------------------------
    // ПРЕДВАРИТЕЛЬНАЯ ФАЗА: сначала полностью проходим МОКАП 2.
    // Для каждого уникального альбома берём первый найденный исходник, открываем
    // 2.psd, вставляем обложку и даём один раз вручную выставить лупу.
    // Положение сохраняется в LOUPE_OFFSETS. После этой фазы основной цикл уже
    // не требует никакого участия пользователя: все остальные расцветки и
    // остальные мокапы собираются автоматически.
    // -------------------------------------------------------------------------
    LOUPE_OFFSETS = {};

    var albumsForLoupe = [];
    var albumRepresentative = {};

    for (var li = 0; li < sources.length; li++) {
        var repBaseName = stripExtension(decodeFileName(sources[li].file.name));
        var repParsed = deriveAlbumAndColor(repBaseName);
        if (!repParsed) continue;

        var repAlbum = sources[li].albumOverride || repParsed.album;
        var repKey = normalizeForCompare(repAlbum);

        if (!albumRepresentative.hasOwnProperty(repKey)) {
            albumRepresentative[repKey] = sources[li];
            albumsForLoupe.push(repKey);
        }
    }

    // Все интерактивные позиции собираются до запуска полной автоматической сборки.
    for (var lp = 0; lp < albumsForLoupe.length; lp++) {
        var loupeKey = albumsForLoupe[lp];
        var loupeSource = albumRepresentative[loupeKey];

        try {
            var loupeDoc = app.open(mockupPaths["2"]);
            try {
                var loupeSO = findLayerRecursive(loupeDoc.layers, CONFIG.soNameSimple);
                if (!loupeSO) {
                    logIssue("Не найден смарт-объект \"" + CONFIG.soNameSimple +
                        "\" в 2.psd — позиция лупы для альбома \"" + loupeKey + "\" не настроена");
                } else {
                    fillSmartObjectWithImage(loupeDoc, loupeSO, loupeSource.file, CONFIG.fitMode);
                    positionLoupeLayer(loupeDoc, loupeKey);
                }
            } finally {
                loupeDoc.close(SaveOptions.DONOTSAVECHANGES);
            }
        } catch (eLoupe) {
            logIssue("Ошибка при предварительной настройке лупы для альбома \"" +
                loupeKey + "\": " + eLoupe.message);
        }
    }

    // После этого места positionLoupeLayer() только применяет уже сохранённые
    // смещения и больше никогда не запускает Free Transform.
    // -------------------------------------------------------------------------

    // typeUnits нужны в POINTS для автоподбора кегля в fitTextLayerToBox,
    // rulerUnits — в PIXELS для расчёта границ смарт-объектов. Сохраняем
    // исходные настройки и восстанавливаем их в конце независимо от исхода.
    var savedRulerUnits = app.preferences.rulerUnits;
    var savedTypeUnits = app.preferences.typeUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.POINTS; // не TypeUnits.PT — такой константы нет

    try {
        for (var s = 0; s < sources.length; s++) {
        var sourceFile = sources[s].file;
        var side = sources[s].side; // "white" | "dark"
        var baseName = stripExtension(decodeFileName(sourceFile.name));

        try {
            var parsed = deriveAlbumAndColor(baseName);
            if (!parsed) {
                logIssue("Не удалось разобрать имя файла \"" + decodeFileName(sourceFile.name) +
                    "\" (нужно минимум два слова: альбом и цвет) — файл пропущен");
                continue;
            }

            // название альбома берём из диалога проверки (показан перед началом обработки),
            // а не пересчитываем — там его можно было поправить вручную
            var albumName = sources[s].albumOverride || parsed.album;
            var albumLoupeKey = normalizeForCompare(albumName);

            var outFolder = ensureFolder(new Folder(outputRoot.fsName + "/" + baseName));

            // ---- мокап 1: текст (Artist + Album Name) + картинка ----
            fillMockup1(mockupPaths["1"], CONFIG.soNameSimple, sourceFile,
                new File(outFolder.fsName + "/1.jpg"),
                artistName, albumName, CONFIG.artistLayerName, CONFIG.albumLayerName);

            // ---- мокап 2: картинка + ручное положение лупы (запоминается по альбому) ----
            fillMockup2(mockupPaths["2"], CONFIG.soNameSimple, sourceFile,
                new File(outFolder.fsName + "/2.jpg"),
                albumLoupeKey);

            // ---- простые мокапы (только картинка, один смарт-объект): 3, 6, 7 ----
            var simpleMockupKeys = ["3", "6", "7"];
            for (var smk = 0; smk < simpleMockupKeys.length; smk++) {
                var mKey = simpleMockupKeys[smk];
                fillSimpleMockup(mockupPaths[mKey], CONFIG.soNameSimple, sourceFile,
                    new File(outFolder.fsName + "/" + mKey + ".jpg"));
            }

            // ---- мокап 4: свой цвет + цвет-партнёр ----
            var doc4 = app.open(mockupPaths["4"]);
            try {
                var ownSOName = (side === "white") ? CONFIG.soNameWhite : CONFIG.soNameDark;
                var otherSOName = (side === "white") ? CONFIG.soNameDark : CONFIG.soNameWhite;
                var ownSO = findLayerRecursive(doc4.layers, ownSOName);
                if (ownSO) fillSmartObjectWithImage(doc4, ownSO, sourceFile, CONFIG.fitMode);
                else logIssue("Не найден смарт-объект \"" + ownSOName + "\" в 4.psd");

                // ищем файл-партнёр той же расцветки в противоположной папке по совпадению альбома
                // (имена файлов декодируем — иначе %20 и т.п. ломают сравнение альбомов)
                var otherFolder = (side === "white") ? darkFolder : whiteFolder;
                var otherFiles = listImageFiles(otherFolder);
                var partnerFile = null;
                for (var p = 0; p < otherFiles.length; p++) {
                    var otherParsed = deriveAlbumAndColor(stripExtension(decodeFileName(otherFiles[p].name)));
                    if (otherParsed && otherParsed.album.toLowerCase() === parsed.album.toLowerCase()) {
                        partnerFile = otherFiles[p];
                        break;
                    }
                }
                if (partnerFile) {
                    var otherSO = findLayerRecursive(doc4.layers, otherSOName);
                    if (otherSO) fillSmartObjectWithImage(doc4, otherSO, partnerFile, CONFIG.fitMode);
                    exportJPEG(doc4, new File(outFolder.fsName + "/4.jpg"), CONFIG.jpegQuality);
                } else {
                    logIssue("Не найден файл-партнёр для \"" + decodeFileName(sourceFile.name) + "\" в папке \"" +
                        decodeFileName(otherFolder.name) + "\" — мокап 4 пропущен для этого исходника");
                }
            } finally {
                doc4.close(SaveOptions.DONOTSAVECHANGES);
            }

            // ---- мокапы 5-3/5-4/5-5: ДРУГИЕ альбомы исполнителя (без текущего) ----
            // Текущий исходник исключается из содержимого "Other Albums" по имени файла
            // (если он там физически лежит) — это делается заново для каждого исходника,
            // поэтому у разных исходников набор "других" и выбранный мокап 5-x может
            // отличаться. Если файлов больше 5 (после исключения текущего), берём только
            // первые 5 (по алфавиту, как их сортирует listImageFiles) и используем мокап 5-5.
            // Если же после исключения текущего осталось ровно 2 "других" — их одних
            // недостаточно для мокапа 5-3 (там 3 места), поэтому текущий исходник
            // подставляется вместе с ними как третья картинка.
            // Примеры: осталось 2 других -> мокап 5-3, вставить эти 2 + текущий.
            //          осталось 3 -> мокап 5-3, вставить все 3 (текущий не участвует).
            //          осталось 4 -> мокап 5-4, вставить все 4.
            //          осталось 5 -> мокап 5-5, вставить все 5.
            //          осталось 6+ -> мокап 5-5, вставить первые 5.
            //          осталось <2 -> мокап 5-х пропущен для этого исходника.
            if (otherAlbumsFolder) {
                var allOtherRaw = listImageFiles(otherAlbumsFolder);
                var currentNameNorm = normalizeForCompare(decodeFileName(sourceFile.name));
                var allOther = [];
                for (var oi = 0; oi < allOtherRaw.length; oi++) {
                    if (normalizeForCompare(decodeFileName(allOtherRaw[oi].name)) !== currentNameNorm) {
                        allOther.push(allOtherRaw[oi]);
                    }
                }

                var filesToPlace = null;
                if (allOther.length < 2) {
                    logIssue("В папке \"" + CONFIG.otherAlbumsFolderName + "\" меньше 2 ДРУГИХ исходников " +
                        "(без учёта текущего: " + allOther.length + ") — мокап 5-х пропущен для " + decodeFileName(sourceFile.name));
                } else if (allOther.length === 2) {
                    // Двух "других" не хватает на мокап 5-3 — добавляем текущий исходник как третью картинку.
                    filesToPlace = [sourceFile].concat(allOther);
                } else {
                    filesToPlace = allOther.slice(0, Math.min(allOther.length, 5));
                }

                if (filesToPlace) {
                    var target = filesToPlace.length;
                    var mockupKey = "5-" + target;

                    if (!mockupPaths[mockupKey]) {
                        logIssue("Нет подходящего мокапа для " + target + " исходников в \"" + CONFIG.otherAlbumsFolderName +
                            "\" — мокап 5-х пропущен для " + decodeFileName(sourceFile.name));
                    } else {
                        var doc5 = app.open(mockupPaths[mockupKey]);
                        try {
                            for (var n = 0; n < filesToPlace.length; n++) {
                                var soN = findLayerRecursive(doc5.layers, CONFIG.soPrefix5x + (n + 1));
                                if (soN) fillSmartObjectWithImage(doc5, soN, filesToPlace[n], CONFIG.fitMode);
                                else logIssue("Не найден смарт-объект \"" + CONFIG.soPrefix5x + (n + 1) + "\" в " + mockupKey + ".psd");
                            }
                            replaceTextSmartObjectWithAutoFit(doc5, CONFIG.artistNameLayerName, artistName, CONFIG.artistNameMaxFontSize);
                            exportJPEG(doc5, new File(outFolder.fsName + "/" + mockupKey + ".jpg"), CONFIG.jpegQuality);
                        } finally {
                            doc5.close(SaveOptions.DONOTSAVECHANGES);
                        }
                    }
                }
            }

            processed++;
        } catch (e) {
            logIssue("Ошибка при обработке \"" + decodeFileName(sourceFile.name) + "\": " + e.message + "\nФайл пропущен, продолжаю.");
        }
        }
    } finally {
        app.preferences.rulerUnits = savedRulerUnits;
        app.preferences.typeUnits = savedTypeUnits;
    }

    // Единственный alert в конце безнадзорного прогона: итог + все накопленные
    // предупреждения (если были) вместо отдельного alert() на каждую проблему.
    var finalMsg = "Готово. Обработано исходников: " + processed + " из " + sources.length + ".";
    if (ISSUES.length > 0) {
        finalMsg += "\n\nПредупреждения (" + ISSUES.length + "):\n- " + ISSUES.join("\n- ");
    }

    alert(finalMsg);
}

main();