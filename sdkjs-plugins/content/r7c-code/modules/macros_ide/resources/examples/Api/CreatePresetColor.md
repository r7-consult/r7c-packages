/**
 * @file CreatePresetColor_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreatePresetColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a color from a preset using `Api.CreatePresetColor()`.
 * It creates a preset color (peachPuff), then uses it along with an RGB color to define gradient stops
 * for a linear gradient fill. Finally, it adds a shape to the active worksheet with this fill and a no-fill stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать цвет из предустановленного набора с помощью `Api.CreatePresetColor()`.
 * Он создает предустановленный цвет (peachPuff), затем использует его вместе с цветом RGB для определения точек градиента
 * для линейной градиентной заливки. Наконец, он добавляет фигуру на активный лист с этой заливкой и обводкой без заливки.
 *
 * @param {string} presetColorName - The name of the preset color to create (e.g., "peachPuff", "lightBlue"). (Имя предустановленного цвета для создания.)
 * @returns {ApiColor} A new ApiColor object representing the preset color. (Новый объект ApiColor, представляющий предустановленный цвет.)
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
        
        // Create a preset color
        // Создание предустановленного цвета
        let presetColor = Api.CreatePresetColor("peachPuff");
        
        // Create gradient stops using the preset color and an RGB color
        // Создание точек градиента с использованием предустановленного цвета и цвета RGB
        let gs1 = Api.CreateGradientStop(presetColor, 0);
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
