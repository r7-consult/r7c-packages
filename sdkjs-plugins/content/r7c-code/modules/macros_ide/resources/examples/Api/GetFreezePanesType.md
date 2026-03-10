/**
 * @file GetFreezePanesType_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetFreezePanesType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the current freeze panes type in a spreadsheet.
 * It first sets the freeze panes type to 'column' for demonstration purposes, then retrieves
 * the current freeze panes type, and displays it in cell B1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текущий тип закрепленных областей в электронной таблице.
 * Сначала он устанавливает тип закрепленных областей на 'column' для демонстрационных целей, затем извлекает
 * текущий тип закрепленных областей и отображает его в ячейке B1 активного листа.
 *
 * @returns {string} The current freeze panes type (e.g., 'column', 'row', 'none'). (Текущий тип закрепленных областей.)
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
        
        // Set freeze panes to freeze the first column for demonstration
        // Установка закрепленных областей для закрепления первого столбца для демонстрации
        Api.SetFreezePanesType('column');
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Display a label in cell A1
        // Отображение метки в ячейке A1
        worksheet.GetRange("A1").SetValue("Type: ");
        
        // Get the current freeze panes type and display it in cell B1
        // Получение текущего типа закрепленных областей и отображение его в ячейке B1
        worksheet.GetRange("B1").SetValue(Api.GetFreezePanesType());
        
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
