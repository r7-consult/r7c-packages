/**
 * @file SetFreezePanesType_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SetFreezePanesType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the freeze panes type in a spreadsheet.
 * It freezes the first column of the active worksheet and then displays the address
 * of the frozen pane in cells A1 and B1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить тип закрепленных областей в электронной таблице.
 * Он закрепляет первый столбец активного листа, а затем отображает адрес
 * закрепленной области в ячейках A1 и B1.
 *
 * @param {string} type - The type of freeze panes to set (e.g., 'column', 'row', 'none'). (Тип закрепленных областей для установки.)
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
        
        // Set freeze panes to freeze the first column
        // Установка закрепленных областей для закрепления первого столбца
        Api.SetFreezePanesType('column');
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Get the freeze panes object and its location
        // Получение объекта закрепленных областей и его местоположения
        let freezePanes = worksheet.GetFreezePanes();
        let range = freezePanes.GetLocation();
        
        // Display the location of the frozen pane in cells A1 and B1
        // Отображение местоположения закрепленной области в ячейках A1 и B1
        worksheet.GetRange("A1").SetValue("Location: ");
        worksheet.GetRange("B1").SetValue(range.GetAddress());
        
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
