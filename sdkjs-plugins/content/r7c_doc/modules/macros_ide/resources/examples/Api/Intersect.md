/**
 * @file Intersect_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.Intersect
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to find the intersection of two or more ranges in a spreadsheet.
 * It defines two ranges, finds their intersection using `Api.Intersect()`, and then fills
 * the intersecting cells with a specific color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как найти пересечение двух или более диапазонов в электронной таблице.
 * Он определяет два диапазона, находит их пересечение с помощью `Api.Intersect()`, а затем заливает
 * пересекающиеся ячейки определенным цветом.
 *
 * @param {ApiRange} range1 - The first range object. (Первый объект диапазона.)
 * @param {ApiRange} range2 - The second range object. (Второй объект диапазона.)
 * @returns {ApiRange} A new ApiRange object representing the intersection of the input ranges. (Новый объект ApiRange, представляющий пересечение входных диапазонов.)
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
        
        // Define two ranges
        // Определение двух диапазонов
        let range1 = worksheet.GetRange("A1:C5");
        let range2 = worksheet.GetRange("B2:B4");
        
        // Find the intersection of the two ranges
        // Нахождение пересечения двух диапазонов
        let intersectionRange = Api.Intersect(range1, range2);
        
        // Fill the intersecting cells with a specific color
        // Заливка пересекающихся ячеек определенным цветом
        intersectionRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
