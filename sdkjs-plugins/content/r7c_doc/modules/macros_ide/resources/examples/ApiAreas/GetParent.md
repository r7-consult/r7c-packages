/**
 * @file GetParent_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiAreas.GetParent
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the parent object for a specified collection of areas.
 * It sets a value in a range, selects it, gets the areas collection, and then retrieves its parent.
 * Finally, it displays information about the parent object and its class type in the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить родительский объект для указанной коллекции областей.
 * Он устанавливает значение в диапазоне, выбирает его, получает коллекцию областей, а затем извлекает ее родителя.
 * Наконец, он отображает информацию о родительском объекте и его типе класса на листе.
 *
 * @returns {object} The parent object of the collection. (Родительский объект коллекции.)
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
        
        // Set a value in range B1:D1 and select it
        // Установка значения в диапазоне B1:D1 и его выбор
        let range = worksheet.GetRange("B1:D1");
        range.SetValue("1");
        range.Select();
        
        // Get the areas collection from the selected range
        // Получение коллекции областей из выбранного диапазона
        let areas = range.GetAreas();
        
        // Get the parent object of the areas collection
        // Получение родительского объекта коллекции областей
        let parent = areas.GetParent();
        
        // Get the class type of the parent object
        // Получение типа класса родительского объекта
        let type = parent.GetClassType();
        
        // Display information about the parent object and its type in the worksheet
        // Отображение информации о родительском объекте и его типе на листе
        range = worksheet.GetRange('A4');
        range.SetValue("The areas parent: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B4').Paste(parent);
        
        range = worksheet.GetRange('A5');
        range.SetValue("The type of the areas parent: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B5').SetValue(type);
        
        // Success notification
        // Уведомление об успешном выполнении
        console.log('Macro executed successfully');
        
    } catch (error) {
        // Error handling
        // Обработка ошибок
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        // Опционально: Показать ошибку пользователю
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
