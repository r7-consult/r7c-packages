/**
 * @file GetType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiProtectedRangeUserInfo.GetType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the type of a protected range user.
 * It adds a protected range to the worksheet, adds a user to it, and then displays the user's type.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип пользователя защищенного диапазона.
 * Он добавляет защищенный диапазон на лист, добавляет в него пользователя, а затем отображает тип пользователя.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Original code enhanced with error handling:
        // This example gets type of a protected range user.
        
        // How to get a user info of a protected range and show its type.
        
        // Get a user type of a protected range and add it to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        let userInfo = protectedRange.GetUser("userId");
        let userType = userInfo.GetType();
        worksheet.GetRange("A3").SetValue("Type: " + userType);
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
