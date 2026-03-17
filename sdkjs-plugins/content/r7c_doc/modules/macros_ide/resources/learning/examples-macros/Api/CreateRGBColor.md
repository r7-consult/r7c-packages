/**
 * @file CreateRGBColor_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateRGBColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create an RGB color by specifying its red, green, and blue components.
 * It then uses this RGB color to create gradient stops for a linear gradient fill,
 * and finally adds a shape to the active worksheet with this fill and a no-fill stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать цвет RGB, указав его красную, зеленую и синюю составляющие.
 * Затем он использует этот цвет RGB для создания градиентных точек для линейной градиентной заливки,
 * и, наконец, добавляет фигуру на активный лист с этой заливкой и обводкой без заливки.
 *
 * @param {number} red - The red component of the color (0-255). (Красная составляющая цвета (0-255).)
 * @param {number} green - The green component of the color (0-255). (Зеленая составляющая цвета (0-255).)
 * @param {number} blue - The blue component of the color (0-255). (Синяя составляющая цвета (0-255).)
 * @returns {ApiColor} A new ApiColor object representing the RGB color. (Новый объект ApiColor, представляющий цвет RGB.)
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
        
        // Create gradient stops using RGB colors
        // Создание градиентных точек с использованием цветов RGB
        let gs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        
        // Create a linear gradient fill
        // Создание линейной градиентной заливки
        let fill = Api.CreateLinearGradientFill([gs1, gs2], 5400000);
        
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
