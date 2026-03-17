/**
 * @file CreateColorByName_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateColorByName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a color by its predefined name.
 * It creates a color using the name "peachPuff" and then applies this color
 * to the font of the text in cell A2 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать цвет по его предопределенному имени.
 * Он создает цвет, используя имя "peachPuff", а затем применяет этот цвет
 * к шрифту текста в ячейке A2 активного листа.
 *
 * @param {string} colorName - The name of the color to create (e.g., "red", "blue", "peachPuff"). (Имя цвета для создания.)
 * @returns {ApiColor} A new ApiColor object representing the named color. (Новый объект ApiColor, представляющий именованный цвет.)
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
        
        // Create a color by its name
        // Создание цвета по его имени
        let color = Api.CreateColorByName("peachPuff");
        
        // Set text in cell A2
        // Установка текста в ячейке A2
        worksheet.GetRange("A2").SetValue("Text with color");
        
        // Set the font color of text in cell A2 using the created named color
        // Установка цвета шрифта текста в ячейке A2 с использованием созданного именованного цвета
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
