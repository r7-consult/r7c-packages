/*********************
 * P7 Офис Spreadsheet Macro Example
 * 
 * Этот макрос демонстрирует основные операции с электронной таблицей:
 * - Получение активного листа
 * - Выбор диапазона ячеек
 * - Настройка цвета фона
 * - Отображение сообщения
 
 -------------
 Lad — центр компетенций по импортозамещению
 
   Помогаем перейти на импортонезависимые программные продукты, 
входящие в реестр российского ПО. 
  Внедряем и сопровождаем отечественное программное обеспечение, 
обучаем пользователей и администраторов 

https://lad-soft.ru/
gov@lad24.ru
--------------

***********************/

(function() {
    // Get the active spreadsheet
    var oSheet = Api.GetActiveSheet();
    if (!oSheet) {
        throw new Error("Active sheet is not available");
    }
    
    // Select range A1:J20
    var oRange = oSheet.GetRange("A1:J20");
    if (!oRange) {
        throw new Error("Range A1:J20 is not available");
    }
    
    // Set background color to red
    var oRed = (Api.CreateColorFromRGB || Api.CreateRGBColor).call(Api, 255, 0, 0);
    oRange.SetFillColor(oRed);
    
    // Show completion message  
    Api.ShowMessage("Spreadsheet Macro", "Cells A1:J20 have been filled with red color.");
    
    // This line will cause an error for demonstration
    // UnknownApi.DoSomething();
    
    console.log("Spreadsheet macro executed successfully!");
})();
