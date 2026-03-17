/**
 * ПРОСТОЙ ТЕСТ Api.SqlNotebook.query()
 *
 * Инструкция:
 * 1. Откройте SQLite Notebook вкладку
 * 2. Загрузите файл MOCK_DATA.csv (через Upload File)
 * 3. Скопируйте код ниже в макрос
 * 4. Нажмите Run (F5)
 * 5. Проверьте результат в ячейках A1-A10
 */

(function() {
    var sheet = Api.GetActiveSheet();

    // Очистка
    sheet.GetRange("A1:B20").Clear();

    sheet.GetRange("A1").SetValue("=== TEST Api.SqlNotebook ===");

    // Проверка 1: Есть ли Api.SqlNotebook?
    if (typeof Api.SqlNotebook === 'undefined') {
        sheet.GetRange("A2").SetValue("ОШИБКА: Api.SqlNotebook не определен!");
        sheet.GetRange("A2").SetFontColor(Api.CreateColorFromRGB(255, 0, 0));
        return;
    }

    sheet.GetRange("A2").SetValue("✓ Api.SqlNotebook найден");
    sheet.GetRange("A2").SetFontColor(Api.CreateColorFromRGB(0, 128, 0));

    // Проверка 2: Есть ли метод query?
    if (typeof Api.SqlNotebook.query !== 'function') {
        sheet.GetRange("A3").SetValue("ОШИБКА: query() не функция!");
        sheet.GetRange("A3").SetFontColor(Api.CreateColorFromRGB(255, 0, 0));
        return;
    }

    sheet.GetRange("A3").SetValue("✓ query() метод найден");
    sheet.GetRange("A3").SetFontColor(Api.CreateColorFromRGB(0, 128, 0));

    // Проверка 3: Выполнить запрос
    sheet.GetRange("A4").SetValue("Выполняю SELECT * FROM data LIMIT 5...");

    var result = Api.SqlNotebook.query(
        'SELECT * FROM "MOCK_DATA.csv" LIMIT 5',
        "MOCK_DATA.csv"
    );

    sheet.GetRange("A5").SetValue("Результат:");

    if (!result.ok) {
        sheet.GetRange("A6").SetValue("ОШИБКА: " + result.error.code);
        sheet.GetRange("A6").SetFontColor(Api.CreateColorFromRGB(255, 0, 0));
        sheet.GetRange("A7").SetValue("Сообщение: " + result.error.message);
        return;
    }

    // Успех!
    sheet.GetRange("A6").SetValue("✓ УСПЕХ! Запрос выполнен");
    sheet.GetRange("A6").SetFontColor(Api.CreateColorFromRGB(0, 128, 0));

    sheet.GetRange("A7").SetValue("Колонок: " + result.result.columnCount);
    sheet.GetRange("A8").SetValue("Строк: " + result.result.rowCount);

    // Вывод колонок
    var cols = result.result.columns.join(", ");
    sheet.GetRange("A9").SetValue("Колонки: " + cols);

    // Вывод первой строки данных
    if (result.result.rows.length > 0) {
        var firstRow = result.result.rows[0].join(", ");
        sheet.GetRange("A10").SetValue("Первая строка: " + firstRow);
    }

    sheet.GetRange("A12").SetValue("=== ТЕСТ ПРОЙДЕН! ===");
    sheet.GetRange("A12").SetFontColor(Api.CreateColorFromRGB(0, 128, 0));
    sheet.GetRange("A12").SetBold(true);
})();
