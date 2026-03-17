---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-format-08"
  title: "Урок 8: Практика. Логическое ветвление и UI Мутации"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        try {
            let sheet = Api.GetActiveSheet();
            
            let redColor   = Api.CreateColorFromRGB(255, 200, 200); // Светло-красный
            let greenColor = Api.CreateColorFromRGB(200, 255, 200); // Светло-зеленый
            
            // ВАША ЗАДАЧА: по каждому KPI проверьте значение и раскрасьте ФОН через SetFillColor(Api.CreateSolidFill(color))
            // 1. C4 Выручка: Норма > 100 -> зелёный. Иначе -> красный.
            let revenue = Number(sheet.GetRange("C4").GetValue());
            if (revenue > 100) {
                sheet.GetRange("C4").SetFillColor(Api.CreateSolidFill(greenColor));
            } else {
                sheet.GetRange("C4").SetFillColor(Api.CreateSolidFill(redColor));
            }
            
            // 2. C5 Отток: Норма < 5 -> зелёный. Иначе -> красный.
            
            
            // 3. C6 NPS: Норма > 70 -> зелёный. Иначе -> красный.
            
            
        } catch (error) {
            console.error(error.message);
        }
    })();
  beforeScript: |
    (function() {
        let sheet = Api.GetActiveSheet();
        sheet.GetRange("A1:Z100").Clear();
        let kpis = [["Выручка (млн)", 120], ["Отток клиентов (%)", 8.5], ["NPS (Лояльность)", 65]];
        sheet.GetRange("B3").SetValue("Показатель");
        sheet.GetRange("C3").SetValue("Значение");
        sheet.GetRange("B3:C3").SetBold(true);
        for(let i=0; i<kpis.length; i++) {
            sheet.GetRange("B" + (i+4)).SetValue(kpis[i][0]);
            sheet.GetRange("C" + (i+4)).SetValue(kpis[i][1]);
        }
        sheet.GetRange("E3:L4").Merge(false);
        sheet.GetRange("E3:L4").SetValue("Монитор KPI CEO: Раскрасьте ФОН значений в C4 (выручка>100 = зелёный), C5 (отток<5 = зелёный), C6 (NPS>70 = зелёный). Иначе красный!");
        sheet.GetRange("E3:L4").SetFontColor(Api.CreateColorFromRGB(100, 150, 100));
    })();
  checks:
      - type: "cell_color"
        cell: "C4"
        expected: "200, 255, 200"
      - type: "cell_color"
        cell: "C5"
        expected: "255, 200, 200"
      - type: "cell_color"
        cell: "C6"
        expected: "255, 200, 200"
---
# Урок 8. Практика: Дашборд KPI Руководителя

Встроенное условное форматирование не справляется с многоуровневой логикой. Программный подход даёт полный контроль.

## Дашборд CEO
* **C4** (Выручка, млн): Норма `> 100`. У нас **120**. -> Зелёный
* **C5** (Отток кл, %): Норма `< 5`. У нас **8.5**. -> Красный
* **C6** (NPS): Норма `> 70`. У нас **65**. -> Красный

Раскрасьте **фон** ячеек C4, C5, C6 методом `.SetFillColor(Api.CreateSolidFill(color))`.

### Готовое решение

```javascript
(function() {
    'use strict';
    try {
        let sheet = Api.GetActiveSheet();
        let redColor   = Api.CreateColorFromRGB(255, 200, 200);
        let greenColor = Api.CreateColorFromRGB(200, 255, 200);
        
        let revenue = Number(sheet.GetRange("C4").GetValue());
        sheet.GetRange("C4").SetFillColor(Api.CreateSolidFill(revenue > 100 ? greenColor : redColor));
        
        let churn = Number(sheet.GetRange("C5").GetValue());
        sheet.GetRange("C5").SetFillColor(Api.CreateSolidFill(churn < 5 ? greenColor : redColor));
        
        let nps = Number(sheet.GetRange("C6").GetValue());
        sheet.GetRange("C6").SetFillColor(Api.CreateSolidFill(nps > 70 ? greenColor : redColor));
    } catch (error) {
        console.error(error.message);
    }
})();
```
