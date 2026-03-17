---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-basics-04-math"
  title: "Урок 4: Практика. Базовая математика и переменные"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        try {
            let sheet = Api.GetActiveSheet();
            
            // ВАША ЗАДАЧА:
            // 1. Считайте выручку за 3 квартала: Q1 (C4), Q2 (F9), Q3 (K15).
            //    ВАЖНО: оберните в Number()! Например: Number(sheet.GetRange("C4").GetValue())
            // 2. Сложите их -> получите "Грязную выручку" (Gross Revenue).
            // 3. Считайте ставку налога из ячейки Z1.
            // 4. Чистая прибыль (Net Profit) = Gross - (Gross * Tax)
            // 5. Запишите результат в ячейку B1.
            
            
            
        } catch (error) {
            console.error(error.message);
        }
    })();
  beforeScript: |
    (function() {
        let sheet = Api.GetActiveSheet();
        sheet.GetRange("A1:Z100").Clear();
        sheet.GetRange("C4").SetValue(1250000);
        sheet.GetRange("C4").SetFillColor(Api.CreateColorFromRGB(235, 245, 255));
        sheet.GetRange("C3").SetValue("Q1 (Каз.)");
        sheet.GetRange("C3").SetBold(true);
        sheet.GetRange("F9").SetValue(840000);
        sheet.GetRange("F9").SetFillColor(Api.CreateColorFromRGB(235, 245, 255));
        sheet.GetRange("F8").SetValue("Q2 (Сиб.)");
        sheet.GetRange("F8").SetBold(true);
        sheet.GetRange("K15").SetValue(2100000);
        sheet.GetRange("K15").SetFillColor(Api.CreateColorFromRGB(235, 245, 255));
        sheet.GetRange("K14").SetValue("Q3 (Мос.)");
        sheet.GetRange("K14").SetBold(true);
        sheet.GetRange("Z1").SetValue(0.15);
        sheet.GetRange("Y1").SetValue("Налог:");
        sheet.GetRange("A1").SetValue("Чистая прибыль:");
        sheet.GetRange("A1").SetBold(true);
        sheet.GetRange("A3:I3").Merge(false);
        sheet.GetRange("A3:I3").SetValue("Финансы: данные за кварталы (C4=Q1, F9=Q2, K15=Q3) и налог (Z1=15%). Выведите чистую прибыль в B1!");
        sheet.GetRange("A3:I3").SetFontColor(Api.CreateColorFromRGB(150, 100, 100));
    })();
  checks:
      - type: "cell_value"
        cell: "B1"
        expected: 3561500
---
# Урок 4. Практика: Финансовая Аналитика

Данные пришли в неструктурированном виде: Q1 в ячейке **C4**, Q2 в **F9**, Q3 аж в **K15**. Налоговая ставка лежит в **Z1**.

## Бизнес-задача
1. `Gross = Q1 + Q2 + Q3`
2. `Net = Gross - (Gross * Tax)`

> **Внимание:** оборачивайте значения из ячеек в `Number()` при математике, иначе JS сложит строки, а не числа!

### Готовое решение

```javascript
(function() {
    'use strict';
    try {
        let sheet = Api.GetActiveSheet();
        
        let q1  = Number(sheet.GetRange("C4").GetValue());
        let q2  = Number(sheet.GetRange("F9").GetValue());
        let q3  = Number(sheet.GetRange("K15").GetValue());
        let tax = Number(sheet.GetRange("Z1").GetValue());
        
        let gross     = q1 + q2 + q3;
        let netProfit = gross - (gross * tax);
        
        sheet.GetRange("B1").SetValue(netProfit);
        
    } catch (error) {
        console.error(error.message);
    }
})();
```
