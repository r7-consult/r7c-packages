/**
 * @file GetCount_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiAreas.GetCount
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro illustrates how to retrieve the number of objects within an `ApiAreas` collection.
 * It sets a value in a range, selects it, obtains the areas collection, and then gets the count
 * of ranges within this collection. Finally, it displays the count in the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить количество объектов в коллекции `ApiAreas`.
 * Он устанавливает значение в диапазоне, выбирает его, получает коллекцию областей, а затем получает количество
 * диапазонов в этой коллекции. Наконец, он отображает количество на листе.
 *
 * @returns {number} The number of objects in the collection. (Количество объектов в коллекции.)
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
        
        // Get the count of items in the areas collection
        // Получение количества элементов в коллекции областей
        let count = areas.GetCount();
        
        // Display the count in the worksheet
        // Отображение количества на листе
        range = worksheet.GetRange('A5');
        range.SetValue("The number of ranges in the areas: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B5').SetValue(count);
        
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
