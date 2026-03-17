/**
 * @file InsertPivotExistingWorksheet_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.InsertPivotExistingWorksheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to insert a pivot table into an existing worksheet.
 * It first sets up sample data in the active worksheet, then defines a data range
 * and a pivot table location, and finally creates a new pivot table in the specified
 * existing worksheet based on this data.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить сводную таблицу в существующий лист.
 * Сначала он настраивает примерные данные на активном листе, затем определяет диапазон данных
 * и местоположение сводной таблицы, и, наконец, создает новую сводную таблицу в указанном
 * существующем листе на основе этих данных.
 *
 * @param {ApiRange} dataRange - The range of data to be used for the pivot table. (Диапазон данных, который будет использоваться для сводной таблицы.)
 * @param {ApiRange} pivotTableLocation - The top-left cell where the pivot table will be inserted. (Левая верхняя ячейка, куда будет вставлена сводная таблица.)
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
        
        // Define the location where the pivot table will be inserted
        // Определение местоположения, куда будет вставлена сводная таблица
        let pivotRef = worksheet.GetRange('A7');
        
        // Insert a new pivot table into the existing worksheet at the specified location
        // Вставка новой сводной таблицы в существующий лист в указанном местоположении
        let pivotTable = Api.InsertPivotExistingWorksheet(dataRef, pivotRef);
        
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
