/**
 * @file GetAllPivotTables_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetAllPivotTables
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve all pivot tables from the active workbook.
 * It first sets up sample data, creates three new pivot tables based on this data,
 * and then iterates through all pivot tables to add a 'Price' data field to each.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все сводные таблицы из активной книги.
 * Сначала он настраивает примерные данные, создает три новые сводные таблицы на основе этих данных,
 * а затем перебирает все сводные таблицы, чтобы добавить поле данных 'Price' к каждой.
 *
 * @returns {Array<ApiPivotTable>} An array of ApiPivotTable objects representing all pivot tables in the workbook. (Массив объектов ApiPivotTable, представляющих все сводные таблицы в книге.)
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        // Инициализация API R7 Office
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Set up sample data for the pivot tables
        // Настройка примерных данных для сводных таблиц
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        // Define the data range for the pivot tables
        // Определение диапазона данных для сводных таблиц
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        
        // Insert three new pivot tables into new worksheets
        // Вставка трех новых сводных таблиц в новые листы
        Api.InsertPivotNewWorksheet(dataRef);
        Api.InsertPivotNewWorksheet(dataRef);
        Api.InsertPivotNewWorksheet(dataRef);
        
        // Get all pivot tables and add a 'Price' data field to each
        // Получение всех сводных таблиц и добавление поля данных 'Price' к каждой
        Api.GetAllPivotTables().forEach(function (pivot) {
            pivot.AddDataField('Price');
        });
        
        // Success notification
        // Уведомление об успешном выполнении
        console.log('Macro executed successfully');
        
    } catch (error) {
        // Error handling
        // Обработка ошибок
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user in cell A1 if API is available
        // Опционально: Показать ошибку пользователю в ячейке A1, если API доступен
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
