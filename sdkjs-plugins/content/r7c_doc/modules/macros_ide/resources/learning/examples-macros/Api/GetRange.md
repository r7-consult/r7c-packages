/**
 * @file GetRange_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetRange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve an `ApiRange` object by its reference.
 * It gets the range "A1:C1", sets its fill color to a specific RGB value,
 * and then displays a message in cell A3 indicating the change.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект `ApiRange` по его ссылке.
 * Он получает диапазон "A1:C1", устанавливает его цвет заливки на определенное значение RGB,
 * а затем отображает сообщение в ячейке A3, указывающее на изменение.
 *
 * @param {string} rangeReference - The reference of the range (e.g., "A1", "B2:D5"). (Ссылка на диапазон.)
 * @returns {ApiRange} The ApiRange object representing the specified range. (Объект ApiRange, представляющий указанный диапазон.)
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
        
        // Get the range A1:C1
        // Получение диапазона A1:C1
        let range = Api.GetRange("A1:C1");
        
        // Set the fill color of the range
        // Установка цвета заливки диапазона
        range.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
        // Display a message indicating the color change
        // Отображение сообщения, указывающего на изменение цвета
        worksheet.GetRange("A3").SetValue("The color was set to the background of cells A1:C1.");
        
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
