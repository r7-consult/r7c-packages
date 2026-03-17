/**
 * @file SetThemeColors_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SetThemeColors
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the theme colors for the current spreadsheet.
 * It retrieves all available theme colors, displays them in column A of the active worksheet,
 * and then applies one of the retrieved themes (the fourth one in this example) to the spreadsheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить цвета темы для текущей электронной таблицы.
 * Он извлекает все доступные цвета темы, отображает их в столбце A активного листа,
 * а затем применяет одну из извлеченных тем (четвертую в этом примере) к электронной таблице.
 *
 * @param {string} themeColor - The theme color to apply. (Цвет темы для применения.)
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
        
        // Get all available theme colors
        // Получение всех доступных цветов темы
        let themes = Api.GetThemesColors();
        
        // Display theme colors in column A
        // Отображение цветов темы в столбце A
        for (let i = 0; i < themes.length; ++i) {
            worksheet.GetRange("A" + (i + 1)).SetValue(themes[i]);
        }
        
        // Set the theme colors to the current spreadsheet (using the 4th theme as an example)
        // Установка цветов темы для текущей электронной таблицы (используя 4-ю тему в качестве примера)
        Api.SetThemeColors(themes[3]);
        
        // Display a message indicating which theme was set
        // Отображение сообщения о том, какая тема была установлена
        worksheet.GetRange("C3").SetValue("The 'Apex' theme colors were set to the current spreadsheet.");
        
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
