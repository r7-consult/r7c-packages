/**
 * ПРОДВИНУТЫЙ ПРИМЕР: SQL запрос + обработка данных
 *
 * Этот макрос показывает:
 * 1. Получение данных через Api.SqlNotebook.query()
 * 2. Вывод в таблицу с форматированием
 * 3. Обработка данных (фильтрация, статистика)
 * 4. Условное форматирование ячеек
 */

(function() {
    var sheet = Api.GetActiveSheet();

    // Очистка области
    sheet.GetRange("A1:H50").Clear();

    // ============================================
    // 1. SQL ЗАПРОС - получаем данные по gender
    // ============================================
    var result = Api.SqlNotebook.query(
        'SELECT first_name, last_name, email, gender FROM "MOCK_DATA.csv" LIMIT 20',
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ОШИБКА: " + result.error.message);
        return;
    }

    var data = result.result;

    // ============================================
    // 2. ВЫВОД ЗАГОЛОВКА
    // ============================================
    sheet.GetRange("A1").SetValue("📊 Отчет по данным из SQL");
    sheet.GetRange("A1").SetBold(true);
    sheet.GetRange("A1").SetFontSize(14);
    sheet.GetRange("A1").SetFontColor(Api.CreateColorFromRGB(0, 100, 200));

    // ============================================
    // 3. ВЫВОД ДАННЫХ С ФОРМАТИРОВАНИЕМ
    // ============================================

    // Заголовки колонок
    var headerRow = 3;
    for (var i = 0; i < data.columns.length; i++) {
        var cell = sheet.GetRange(String.fromCharCode(65 + i) + headerRow);
        cell.SetValue(data.columns[i]);
        cell.SetBold(true);
        cell.SetFillColor(Api.CreateColorFromRGB(200, 200, 200));
        cell.SetAlignHorizontal("center");
    }

    // Данные
    var startRow = headerRow + 1;
    for (var r = 0; r < data.rows.length; r++) {
        for (var c = 0; c < data.rows[r].length; c++) {
            var cell = sheet.GetRange(String.fromCharCode(65 + c) + (startRow + r));
            cell.SetValue(data.rows[r][c]);

            // Условное форматирование по gender (колонка 3)
            if (c === 3) { // gender column
                if (data.rows[r][c] === "Male") {
                    cell.SetFillColor(Api.CreateColorFromRGB(200, 230, 255)); // Голубой
                } else if (data.rows[r][c] === "Female") {
                    cell.SetFillColor(Api.CreateColorFromRGB(255, 200, 230)); // Розовый
                }
            }
        }
    }

    // ============================================
    // 4. СТАТИСТИКА - обработка данных в JS
    // ============================================

    // Подсчет Male/Female
    var maleCount = 0;
    var femaleCount = 0;
    var genderColumnIndex = 3; // gender колонка

    for (var i = 0; i < data.rows.length; i++) {
        var gender = data.rows[i][genderColumnIndex];
        if (gender === "Male") maleCount++;
        else if (gender === "Female") femaleCount++;
    }

    // Вывод статистики
    var statsRow = startRow + data.rows.length + 2;

    sheet.GetRange("A" + statsRow).SetValue("📈 Статистика:");
    sheet.GetRange("A" + statsRow).SetBold(true);

    sheet.GetRange("A" + (statsRow + 1)).SetValue("Всего записей:");
    sheet.GetRange("B" + (statsRow + 1)).SetValue(data.rowCount);

    sheet.GetRange("A" + (statsRow + 2)).SetValue("Мужчин:");
    sheet.GetRange("B" + (statsRow + 2)).SetValue(maleCount);
    sheet.GetRange("B" + (statsRow + 2)).SetFillColor(Api.CreateColorFromRGB(200, 230, 255));

    sheet.GetRange("A" + (statsRow + 3)).SetValue("Женщин:");
    sheet.GetRange("B" + (statsRow + 3)).SetValue(femaleCount);
    sheet.GetRange("B" + (statsRow + 3)).SetFillColor(Api.CreateColorFromRGB(255, 200, 230));

    // Процентное соотношение
    var malePercent = Math.round((maleCount / data.rowCount) * 100);
    var femalePercent = Math.round((femaleCount / data.rowCount) * 100);

    sheet.GetRange("A" + (statsRow + 4)).SetValue("Процент мужчин:");
    sheet.GetRange("B" + (statsRow + 4)).SetValue(malePercent + "%");

    sheet.GetRange("A" + (statsRow + 5)).SetValue("Процент женщин:");
    sheet.GetRange("B" + (statsRow + 5)).SetValue(femalePercent + "%");

    // ============================================
    // 5. ВТОРОЙ SQL ЗАПРОС - агрегация
    // ============================================

    var statsResult = Api.SqlNotebook.query(
        'SELECT gender, COUNT(*) as count FROM "MOCK_DATA.csv" GROUP BY gender',
        "MOCK_DATA.csv"
    );

    if (statsResult.ok) {
        var statsStartRow = statsRow + 7;

        sheet.GetRange("D" + statsStartRow).SetValue("📊 SQL Агрегация:");
        sheet.GetRange("D" + statsStartRow).SetBold(true);

        // Вывод результата GROUP BY
        sheet.GetRange("D" + (statsStartRow + 1)).SetValue("Gender");
        sheet.GetRange("E" + (statsStartRow + 1)).SetValue("Count");
        sheet.GetRange("D" + (statsStartRow + 1) + ":E" + (statsStartRow + 1)).SetBold(true);

        for (var i = 0; i < statsResult.result.rows.length; i++) {
            sheet.GetRange("D" + (statsStartRow + 2 + i)).SetValue(statsResult.result.rows[i][0]);
            sheet.GetRange("E" + (statsStartRow + 2 + i)).SetValue(statsResult.result.rows[i][1]);
        }
    }

    // ============================================
    // 6. ФИНАЛЬНОЕ СООБЩЕНИЕ
    // ============================================

    sheet.GetRange("A" + (statsRow + 10)).SetValue("✅ Готово! Данные загружены и обработаны.");
    sheet.GetRange("A" + (statsRow + 10)).SetFontColor(Api.CreateColorFromRGB(0, 150, 0));
    sheet.GetRange("A" + (statsRow + 10)).SetBold(true);

})();
