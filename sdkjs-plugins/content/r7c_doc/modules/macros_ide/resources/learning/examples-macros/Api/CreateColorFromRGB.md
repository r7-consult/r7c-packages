/**
 * @file CreateColorFromRGB_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateColorFromRGB
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create an RGB color by specifying its red, green, and blue components.
 * It then uses this custom RGB color to set the font color of text in cell A2 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать цвет RGB, указав его красную, зеленую и синюю составляющие.
 * Затем он использует этот пользовательский цвет RGB для установки цвета шрифта текста в ячейке A2 активного листа.
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
        
        // Create an RGB color
        // Создание цвета RGB
        let color = Api.CreateColorFromRGB(255, 111, 61);
        
        // Set text in cell A2
        // Установка текста в ячейке A2
        worksheet.GetRange("A2").SetValue("Text with color");
        
        // Set the font color of text in cell A2 using the created RGB color
        // Установка цвета шрифта текста в ячейке A2 с использованием созданного цвета RGB
        worksheet.GetRange("A2").SetFontColor(color);
        
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
