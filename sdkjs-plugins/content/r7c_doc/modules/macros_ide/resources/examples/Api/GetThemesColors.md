/**
 * @file GetThemesColors_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetThemesColors
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve all available theme colors for the spreadsheet.
 * It gets a list of all theme colors and then displays each color's name in column A
 * of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все доступные цвета темы для электронной таблицы.
 * Он получает список всех цветов темы, а затем отображает название каждого цвета в столбце A
 * активного листа.
 *
 * @returns {Array<string>} An array of strings, where each string represents a theme color. (Массив строк, где каждая строка представляет цвет темы.)
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
        
        // Get all available theme colors
        // Получение всех доступных цветов темы
        let themes = Api.GetThemesColors();
        
        // Display theme colors in column A
        // Отображение цветов темы в столбце A
        for (let i = 0; i < themes.length; ++i) {
            worksheet.GetRange("A" + (i + 1)).SetValue(themes[i]);
        }
        
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
