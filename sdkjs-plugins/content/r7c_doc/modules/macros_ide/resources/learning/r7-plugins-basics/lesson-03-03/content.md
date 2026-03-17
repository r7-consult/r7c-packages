**Урок 3.3. Отладка через консоль Chromium и прямая правка установленных файлов без переустановки**

Для эффективной отладки и доработки плагина **не требуется повторная упаковка и переустановка**. Достаточно редактировать файлы напрямую в папке установленного плагина.

**Шаг 1. Запуск Р7-Офис с поддержкой отладки**  
Запустите Р7-Офис из командной строки с флагом:

R7Office.exe --ascdesktop-support-debug-info (или создайте ярлык с этим параметром)

**Шаг 2. Вызов консоли отладки**

1.  Запустите плагин.
2.  Щёлкните **правой кнопкой мыши** в любом свободном месте окна плагина.
3.  В контекстном меню выберите **«Inspect»** (или нажмите **F1**).
4.  Откроется **DevTools Chromium**.

**Шаг 3. Работа с DevTools**

-    Во вкладке **Console** отображаются ошибки JavaScript.
-    Во вкладке **Sources** можно найти ваш файл code.js:
    -    Перейдите в дерево: frameEditor(...) → pluginFrameEditor(...) → file: → code.js
    -    Установите точки останова (breakpoints), просматривайте значения переменных.
-    Изменения в коде через DevTools **временны** — для постоянного результата редактируйте исходные файлы.

**Шаг 4. Прямая правка файлов без переустановки**  
Установленные плагины хранятся в следующих папках:

-    **Windows**:  
    C:\\Users\\<Имя\_пользователя>\\AppData\\Local\\R7-Office\\Editors\\data\\sdkjs-plugins\\\[GUID\]
-    **Linux**:  
    /home/<пользователь>/.local/share/r7-office/editors/sdkjs-plugins/\[GUID\]

**Процедура доработки**:

1.  Откройте нужный файл (например, code.js) в вашем редакторе кода.
2.  Внесите изменения и сохраните.
3.  **Закройте окно плагина** в Р7-Офис.
4.  **Повторно запустите плагин** — изменения применятся автоматически.

Обратите внимание:

-    При изменении config.json может потребоваться **полная переустановка** плагина.
-    После каждой правки логики (code.js, index.html) достаточно просто **перезапустить плагин**, не перезагружая весь редактор.

Этот модуль завершает цикл разработки первого плагина: от идеи — до рабочего инструмента с возможностью быстрой итеративной доработки.

**Практическая работа №1: Создание плагина «Окрашивание ячейки»**

**Цель**

Реализовать простой плагин, который при нажатии кнопки изменяет цвет активной ячейки на серый.

**Исходные требования**

-    Установлен табличный редактор **Р7-Офис**
-    Плагин разрабатывается во внешнем редакторе кода (например, VS Code, Notepad++)
-    Редактор запущен с флагом --ascdesktop-support-debug-info (для отладки)

**Этапы выполнения**

1.  **Создайте рабочую папку** для проекта (например, cell-color-plugin).
2.  **Создайте файл config.json** со следующим содержимым:

json

{

"name": "Плагин окрашивания ячейки",

"guid": "asc.{ВАШ\_УНИКАЛЬНЫЙ\_GUID}",

"version": "0.0.1",

"variations": \[

{

"description": "Учебный плагин",

"url": "index.html",

"icons": \[\],

"isViewer": true,

"EditorsSupport": \["cell"\],

"isVisual": true,

"isModal": false,

"isInsideMode": true,

"buttons": \[\]

}

\]

}

Замените ВАШ\_УНИКАЛЬНЫЙ\_GUID на GUID, сгенерированный на сайте [https://www.guidgenerator.com/](https://www.guidgenerator.com/).

1.  **Создайте файл index.html**:

Html

<!DOCTYPE html>

<html lang="ru">

<head>

<meta charset="UTF-8" />

<script src="code.js"></script>

</head>

<body>

<button id="btn-1">Сменить цвет</button>

</body>

</html>

1.  **Создайте файл code.js**:

Javascript

(function(window, undefined) {

window.Asc.plugin.init = function() {

document.getElementById('btn-1').addEventListener('click', changeColor);

};

window.Asc.plugin.button = function(id) {

this.executeCommand("close", "");

};

function changeColor() {

window.Asc.plugin.callCommand(function() {

let oWorksheet = Api.GetActiveSheet();

if (oWorksheet) {

let oCell = oWorksheet.GetActiveCell();

if (oCell) {

oCell.SetFillColor(Api.CreateColorFromRGB(240, 240, 240));

}

}

}, undefined, true);

}

})(window, undefined);

1.  **Упакуйте файлы в архив**:
    -    Выделите все три файла (config.json, index.html, code.js)
    -    Создайте ZIP-архив → переименуйте расширение на .plugin (например, cell-color.plugin)
2.  **Установите плагин в Р7-Офис**:
    -    Откройте Р7 → вкладка **«Плагины»** → **«Установить плагин»**
    -    Выберите ваш .plugin-файл
3.  **Протестируйте работу**:
    -    Откройте новый лист
    -    Выделите любую ячейку
    -    Запустите плагин и нажмите кнопку **«Сменить цвет»**
    -    Убедитесь, что ячейка стала серой

**Практическая работа №2: Доработка плагина — добавление выбора цвета**

**Цель**

Расширить функционал плагина: добавить возможность выбора цвета из предложенного списка.

**Этапы выполнения**

1.  **Откройте установленную папку плагина**:
    -    Windows:  
        C:\\Users\\<Ваш\_пользователь>\\AppData\\Local\\R7-Office\\Editors\\data\\sdkjs-plugins\\\[ваш-GUID\]
    -    Откройте файлы index.html и code.js в редакторе кода
2.  **Обновите index.html**:

html

<!DOCTYPE html>

<html lang="ru">

<head>

<meta charset="UTF-8" />

<script src="code.js"></script>

</head>

<body>

<label>Выберите цвет:</label><br>

<select id="color-select">

<option value="240,240,240">Серый</option>

<option value="255,255,0">Жёлтый</option>

<option value="144,238,144">Светло-зелёный</option>

<option value="221,160,221">Сиреневый</option>

</select><br><br>

<button id="btn-apply">Применить</button>

</body>

</html>

1.  **Обновите code.js**:

Javascript

(function(window, undefined) {

window.Asc.plugin.init = function() {

document.getElementById('btn-apply').addEventListener('click', applyColor);

};

window.Asc.plugin.button = function(id) {

this.executeCommand("close", "");

};

function applyColor() {

const colorStr = document.getElementById('color-select').value;

const \[r, g, b\] = colorStr.split(',').map(Number);

window.Asc.plugin.callCommand(function() {

let oWorksheet = Api.GetActiveSheet();

if (oWorksheet) {

let oCell = oWorksheet.GetActiveCell();

if (oCell) {

oCell.SetFillColor(Api.CreateColorFromRGB(r, g, b));

}

}

}, undefined, true);

}

})(window, undefined);

1.  **Сохраните изменения** в оба файла.
2.  **Перезапустите плагин**:
    -    Закройте окно плагина в Р7
    -    Снова нажмите на его кнопку в ленте
3.  **Протестируйте**:
    -    Выберите любой цвет из выпадающего списка
    -    Нажмите **«Применить»**
    -    Убедитесь, что активная ячейка окрасилась в выбранный цвет

**Совет**: теперь вы можете легко добавлять новые цвета, редактируя только HTML-файл — без переустановки!