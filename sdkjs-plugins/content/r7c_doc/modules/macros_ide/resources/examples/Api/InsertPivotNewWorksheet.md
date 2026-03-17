/**
 * @file InsertPivotNewWorksheet_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.InsertPivotNewWorksheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to insert a pivot table into a new worksheet.
 * It first sets up sample data in the active worksheet, then defines a data range,
 * and finally creates a new pivot table in a new worksheet based on this data.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить сводную таблицу в новый лист.
 * Сначала он настраивает примерные данные на активном листе, затем определяет диапазон данных,
 * и, наконец, создает новую сводную таблицу на новом листе на основе этих данных.
 *
 * @param {ApiRange} dataRange - The range of data to be used for the pivot table. (Диапазон данных, который будет использоваться для сводной таблицы.)
 * @returns {ApiPivotTable} The newly created pivot table object. (Вновь созданный объект сводной таблицы.)
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
        
        // Set up sample data for the pivot table
        // Настройка примерных данных для сводной таблицы
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        // Define the data range for the pivot table
        // Определение диапазона данных для сводной таблицы
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        
        // Insert a new pivot table into a new worksheet
        // Вставка новой сводной таблицы в новый лист
        let pivotTable = Api.InsertPivotNewWorksheet(dataRef);
        
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
