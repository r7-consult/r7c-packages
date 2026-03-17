/**
 * @file GetSheets_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetSheets
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve all sheets in the active workbook.
 * It adds a new sheet, then gets a collection of all sheets, and displays the names
 * of the first two sheets in cells A1 and A2 of the newly added sheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все листы в активной книге.
 * Он добавляет новый лист, затем получает коллекцию всех листов и отображает имена
 * первых двух листов в ячейках A1 и A2 вновь добавленного листа.
 *
 * @returns {Array<ApiWorksheet>} An array of ApiWorksheet objects representing all sheets in the workbook. (Массив объектов ApiWorksheet, представляющих все листы в книге.)
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
        
        // Add a new sheet for demonstration purposes
        // Добавление нового листа для демонстрационных целей
        Api.AddSheet("new_sheet_name");
        
        // Get all sheets in the workbook
        // Получение всех листов в книге
        let sheets = Api.GetSheets();
        
        // Get names of the first two sheets
        // Получение имен первых двух листов
        let sheetName1 = sheets[0].GetName();
        let sheetName2 = sheets[1].GetName();
        
        // Display sheet names in the newly added sheet
        // Отображение имен листов во вновь добавленном листе
        sheets[1].GetRange("A1").SetValue(sheetName1);
        sheets[1].GetRange("A2").SetValue(sheetName2);
        
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
