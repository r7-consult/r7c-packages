/**
 * @file RefreshAllPivots_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.RefreshAllPivots
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to refresh all pivot tables in the active workbook.
 * It first sets up sample data, creates a new pivot table based on this data,
 * adds fields to the pivot table, and then calls `Api.RefreshAllPivots()` to update it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как обновить все сводные таблицы в активной книге.
 * Сначала он настраивает примерные данные, создает новую сводную таблицу на основе этих данных,
 * добавляет поля в сводную таблицу, а затем вызывает `Api.RefreshAllPivots()` для ее обновления.
 *
 * @returns {void}
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
        
        // Add fields to the pivot table
        // Добавление полей в сводную таблицу
        Api.GetPivotByName(pivotTable.GetName()).AddFields({
            rows: 'Region',
        });
        
        // Add data field to the pivot table
        // Добавление поля данных в сводную таблицу
        Api.GetPivotByName(pivotTable.GetName()).AddDataField('Price');
        
        // Refresh all pivot tables in the workbook
        // Обновление всех сводных таблиц в книге
        Api.RefreshAllPivots();
        
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
