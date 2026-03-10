/**
 * @file SqlNotebook_Query_Test_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SqlNotebook.query() Test
 * @author R7-Consult
 * @version 1.0.0
 * @date February 2, 2026
 *
 * @description
 * This macro demonstrates how to use the new Api.SqlNotebook.query() method
 * to execute SQL queries against files loaded in SQLite Notebook.
 *
 * Prerequisites:
 * 1. Open SQLite Notebook tab
 * 2. Load a CSV file (e.g., "MOCK_DATA.csv")
 * 3. Run this macro
 *
 * @description (Russian)
 * Этот макрос демонстрирует использование нового метода Api.SqlNotebook.query()
 * для выполнения SQL-запросов к файлам, загруженным в SQLite Notebook.
 *
 * Требования:
 * 1. Открыть вкладку SQLite Notebook
 * 2. Загрузить CSV файл (например, "MOCK_DATA.csv")
 * 3. Запустить этот макрос
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';

    try {
        // Имя файла для тестирования (замените на ваш файл)
        var filename = "MOCK_DATA.csv";

        // Тест 1: Простой SELECT запрос
        console.log("=== Тест 1: Простой SELECT запрос ===");
        var result1 = Api.SqlNotebook.query(
            "SELECT * FROM data LIMIT 5",
            filename
        );

        console.log("Результат теста 1:", result1);

        if (!result1.ok) {
            Api.ShowMessage("Тест 1 провален", result1.error.message);
            return;
        }

        // Вывод результата в консоль
        console.log("Найдено строк:", result1.result.rowCount);
        console.log("Столбцов:", result1.result.columnCount);
        console.log("Колонки:", result1.result.columns);

        // Тест 2: Агрегационный запрос
        console.log("\n=== Тест 2: Агрегационный запрос ===");
        var result2 = Api.SqlNotebook.query(
            "SELECT COUNT(*) as total FROM data",
            filename
        );

        console.log("Результат теста 2:", result2);

        if (!result2.ok) {
            Api.ShowMessage("Тест 2 провален", result2.error.message);
            return;
        }

        console.log("Всего записей:", result2.result.rows[0][0]);

        // Тест 3: Запрос с фильтрацией (замените условие на подходящее для вашего файла)
        console.log("\n=== Тест 3: Запрос с фильтрацией ===");
        var result3 = Api.SqlNotebook.query(
            "SELECT * FROM data LIMIT 10",
            filename
        );

        console.log("Результат теста 3:", result3);

        if (!result3.ok) {
            Api.ShowMessage("Тест 3 провален", result3.error.message);
            return;
        }

        // Тест 4: Вывод результата в таблицу Excel
        console.log("\n=== Тест 4: Вывод в таблицу Excel ===");
        var result4 = Api.SqlNotebook.query(
            "SELECT * FROM data LIMIT 20",
            filename
        );

        if (!result4.ok) {
            Api.ShowMessage("Тест 4 провален", result4.error.message);
            return;
        }

        // Получение активного листа
        var sheet = Api.GetActiveSheet();
        var df = result4.result;

        // Очистка области (опционально)
        var clearRange = sheet.GetRange("A1:Z100");
        clearRange.Clear();

        // Вывод заголовков
        df.columns.forEach(function(col, i) {
            var cell = String.fromCharCode(65 + i) + '1';
            sheet.GetRange(cell).SetValue(col);
            // Форматирование заголовков (жирный шрифт)
            sheet.GetRange(cell).SetBold(true);
        });

        // Вывод данных
        df.rows.forEach(function(row, rowIdx) {
            row.forEach(function(val, colIdx) {
                var cell = String.fromCharCode(65 + colIdx) + (rowIdx + 2);
                sheet.GetRange(cell).SetValue(val);
            });
        });

        // Автоподбор ширины колонок
        var lastColumn = String.fromCharCode(65 + df.columnCount - 1);
        var autoFitRange = sheet.GetRange("A1:" + lastColumn + "1");
        autoFitRange.AutoFit(false, true);

        // Тест 5: Обработка ошибки - несуществующий файл
        console.log("\n=== Тест 5: Обработка ошибки (несуществующий файл) ===");
        var result5 = Api.SqlNotebook.query(
            "SELECT * FROM data",
            "NON_EXISTENT_FILE.csv"
        );

        console.log("Результат теста 5 (должна быть ошибка):", result5);

        if (result5.ok) {
            console.error("ОШИБКА: Тест 5 должен был вернуть ошибку!");
        } else {
            console.log("Корректно обработана ошибка:", result5.error.code, "-", result5.error.message);
        }

        // Тест 6: Обработка ошибки - невалидный SQL
        console.log("\n=== Тест 6: Обработка ошибки (невалидный SQL) ===");
        var result6 = Api.SqlNotebook.query(
            "INVALID SQL SYNTAX",
            filename
        );

        console.log("Результат теста 6 (должна быть ошибка):", result6);

        if (result6.ok) {
            console.error("ОШИБКА: Тест 6 должен был вернуть ошибку!");
        } else {
            console.log("Корректно обработана ошибка:", result6.error.code, "-", result6.error.message);
        }

        // Успешное завершение
        Api.ShowMessage(
            "Тесты завершены",
            "Все тесты выполнены успешно!\n\n" +
            "Проверьте консоль браузера для деталей.\n" +
            "Результаты выведены на активный лист Excel.\n\n" +
            "Загружено строк: " + df.rowCount
        );

        console.log("\n=== Все тесты завершены успешно ===");

    } catch (error) {
        // Обработка ошибок
        console.error('Ошибка выполнения макроса:', error.message);
        Api.ShowMessage('Ошибка', 'Произошла ошибка: ' + error.message);

        // Опционально: Показать ошибку в ячейке A1
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            var sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
