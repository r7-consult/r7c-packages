/**
 * ОТЛАДОЧНЫЙ ТЕСТ - проверка по шагам
 */

(function() {
    var sheet = Api.GetActiveSheet();

    // Шаг 1: Проверка что Api работает
    sheet.GetRange("A1").SetValue("Шаг 1: Api работает");

    // Шаг 2: Проверка что Api.SqlNotebook существует
    if (typeof Api.SqlNotebook === 'undefined') {
        sheet.GetRange("A2").SetValue("ОШИБКА: Api.SqlNotebook не определен!");
        return;
    }
    sheet.GetRange("A2").SetValue("Шаг 2: Api.SqlNotebook существует");

    // Шаг 3: Проверка что query метод существует
    if (typeof Api.SqlNotebook.query !== 'function') {
        sheet.GetRange("A3").SetValue("ОШИБКА: Api.SqlNotebook.query не функция!");
        return;
    }
    sheet.GetRange("A3").SetValue("Шаг 3: query метод существует");

    // Шаг 4: Попытка выполнить запрос
    var result = Api.SqlNotebook.query('SELECT * FROM "MOCK_DATA.csv" LIMIT 3', "MOCK_DATA.csv");

    // Шаг 5: Проверка результата
    sheet.GetRange("A5").SetValue("Результат:");
    sheet.GetRange("B5").SetValue(JSON.stringify(result));

    if (!result.ok) {
        sheet.GetRange("A6").SetValue("ОШИБКА:");
        sheet.GetRange("A7").SetValue("Код: " + result.error.code);
        sheet.GetRange("A8").SetValue("Сообщение: " + result.error.message);
        return;
    }

    // Шаг 6: Вывод данных
    sheet.GetRange("A10").SetValue("Успех! Данные получены:");
    sheet.GetRange("A11").SetValue("Строк: " + result.result.rowCount);
    sheet.GetRange("A12").SetValue("Колонок: " + result.result.columnCount);

    // Шаг 7: Вывод первых 3 строк
    var df = result.result;

    // Заголовки
    sheet.GetRange("A14").SetValue(df.columns[0] || "col0");
    sheet.GetRange("B14").SetValue(df.columns[1] || "col1");
    sheet.GetRange("C14").SetValue(df.columns[2] || "col2");

    // Данные
    if (df.rows.length > 0) {
        sheet.GetRange("A15").SetValue(df.rows[0][0]);
        sheet.GetRange("B15").SetValue(df.rows[0][1]);
        sheet.GetRange("C15").SetValue(df.rows[0][2]);
    }

    if (df.rows.length > 1) {
        sheet.GetRange("A16").SetValue(df.rows[1][0]);
        sheet.GetRange("B16").SetValue(df.rows[1][1]);
        sheet.GetRange("C16").SetValue(df.rows[1][2]);
    }

    sheet.GetRange("A18").SetValue("ТЕСТ ЗАВЕРШЕН УСПЕШНО!");
})();
