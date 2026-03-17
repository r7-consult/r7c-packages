/**
 * @file CreatePatternFill_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreatePatternFill
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a pattern fill to apply to an object.
 * It creates a pattern fill with a specified pattern type ("dashDnDiag"), a foreground color,
 * and a background color. Finally, it adds a shape to the active worksheet with this pattern fill
 * and a no-fill stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать узорную заливку для применения к объекту.
 * Он создает узорную заливку с указанным типом узора ("dashDnDiag"), цветом переднего плана
 * и цветом фона. Наконец, он добавляет фигуру на активный лист с этой узорной заливкой
 * и обводкой без заливки.
 *
 * @param {string} patternType - The type of pattern to use (e.g., "dashDnDiag", "solid"). (Тип узора для использования.)
 * @param {ApiColor} foregroundColor - The foreground color of the pattern. (Цвет переднего плана узора.)
 * @param {ApiColor} backgroundColor - The background color of the pattern. (Цвет фона узора.)
 * @returns {ApiFill} A new ApiFill object representing a pattern fill. (Новый объект ApiFill, представляющий узорную заливку.)
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
        
        // Create a pattern fill with specified pattern type, foreground, and background colors
        // Создание узорной заливки с указанным типом узора, цветами переднего плана и фона
        let fill = Api.CreatePatternFill("dashDnDiag", Api.CreateRGBColor(255, 111, 61), Api.CreateRGBColor(51, 51, 51));
        
        // Create a no-fill stroke
        // Создание обводки без заливки
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a shape to the worksheet with the defined fill and stroke
        // Добавление фигуры на лист с определенной заливкой и обводкой
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        
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
