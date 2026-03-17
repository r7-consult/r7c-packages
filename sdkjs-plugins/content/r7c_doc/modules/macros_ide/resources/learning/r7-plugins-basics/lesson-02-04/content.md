**Урок 2.4. Написание JavaScript-логики: window.Asc.plugin.init() и window.Asc.plugin.button()**

Основная логика плагина реализуется в JavaScript-файле (например, code.js). Ключевыми являются две функции, которые **обязательно** должен содержать каждый плагин:

1.  **window.Asc.plugin.init()** — вызывается при запуске плагина. Здесь регистрируются обработчики событий (например, кликов по кнопкам).
2.  **window.Asc.plugin.button(id)** — вызывается при закрытии модального окна или нажатии системной кнопки. Обычно используется для завершения работы.

Пример code.js:

javascript

(function(window, undefined) {

// Инициализация плагина

window.Asc.plugin.init = function() {

document.getElementById('btn-1').addEventListener('click', changeColor);

};

// Обработка нажатия кнопки в интерфейсе

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

// Завершение работы плагина

window.Asc.plugin.button = function(id) {

this.executeCommand("close", "");

};

})(window, undefined);

**Важно:**

-    Вся работа с документом (чтение/запись ячеек) выполняется внутри callCommand().
-    Параметр true в конце callCommand(..., true) обязателен — он заставляет редактор перерисовать лист после изменений.
-    Функция Api предоставляет доступ к объектной модели Р7 (листы, ячейки, форматирование).

Без корректной реализации init() плагин запустится, но не будет реагировать на действия пользователя.