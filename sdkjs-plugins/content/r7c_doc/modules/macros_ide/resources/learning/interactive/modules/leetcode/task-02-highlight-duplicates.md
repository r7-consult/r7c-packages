---
interactive:
  version: 1
  enabled: true
  lessonId: "leetcode-002-highlight-duplicates"
  title: "Поиск и выделение дубликатов"
  targetEditor: "cell"
  difficulty: "medium"
  task:
    goal: "Напишите макрос для поиска дублирующихся ID пользователей и заливки их ячеек красным."
    steps:
      - "Получите активный лист и диапазон ячеек A2:A7"
      - "Создайте словарь (объект JS) для подсчета количества каждого ID"
      - "Пройдите по диапазону еще раз и, если ID встречается больше одного раза, выделите эту ячейку цветом '#FF0000'"
      - "Для выделения цветом используйте метод .SetFillColor()"
    acceptance:
      - "Ячейки A2, A4, A5 выделены красным, так как содержат дубликаты (ID 101 и 103)."
      - "Ячейки A3, A6, A7 имеют стандартный или прозрачный фон."
  starterCode: |
    (function () {
      const ws = Api.GetActiveSheet();
      // Напишите макрос для поиска дубликатов в диапазоне A2:A7
      // Если ID повторяется больше одного раза в таблице, залейте ячейку красным цветом (Api.CreateColorFromHex("#FF0000"))
      
    })();
  beforeScript: |
    (function () {
      const ws = Api.GetActiveSheet();
      ws.GetRange("A1:C20").Clear();
      
      // Заголовки
      ws.GetRange("A1").SetValue("User ID");
      ws.GetRange("A1").SetBold(true);

      // Данные
      ws.GetRange("A2").SetValue("101");
      ws.GetRange("A3").SetValue("102");
      ws.GetRange("A4").SetValue("103");
      ws.GetRange("A5").SetValue("101");
      ws.GetRange("A6").SetValue("104");
      ws.GetRange("A7").SetValue("103");
      
      // Снимаем заливку для чистоты
      ws.GetRange("A2:A7").SetFillColor(Api.CreateColorFromHex("#FFFFFF"));
    })();
  checks:
    - id: "check-color-a2"
      type: "cell_color"
      sheet: "Active"
      cell: "A2"
      expected: "#FF0000"
    - id: "check-color-a3"
      type: "cell_color"
      sheet: "Active"
      cell: "A3"
      expected: "#FFFFFF"
    - id: "check-color-a4"
      type: "cell_color"
      sheet: "Active"
      cell: "A4"
      expected: "#FF0000"
    - id: "check-color-a5"
      type: "cell_color"
      sheet: "Active"
      cell: "A5"
      expected: "#FF0000"
    - id: "check-color-a6"
      type: "cell_color"
      sheet: "Active"
      cell: "A6"
      expected: "#FFFFFF"
    - id: "check-color-a7"
      type: "cell_color"
      sheet: "Active"
      cell: "A7"
      expected: "#FF0000"
  hints:
    - "Сначала прочитайте весь диапазон: const data = ws.GetRange('A2:A7').GetValue();"
    - "Создайте счетчик: const counts = {}; data.forEach(row => counts[row[0]] = (counts[row[0]] || 0) + 1);"
    - "Вторым циклом пройдитесь по строкам. Внимание: индексы в GetRange() начинаются с 1!"
    - "Для первой строки данных (A2) индекс равен 2: ws.GetRange('A2')"
    - "Чтобы установить цвет: cell.SetFillColor(Api.CreateColorFromHex('#FF0000'));"
  source:
    category: "Практика"
    examplePath: ""
  draft: false
---
# Выделение дубликатов

Вы работаете с базой данных клиентов, и к вам попал список User ID, в котором возможны дубли. Чтобы визуально найти ошибки, вам необходимо написать макрос, который подсветит ячейки с дублирующимися ID **красным цветом**.

## Условия
Диапазон данных — `A2:A7`.

**Вам нужно:**
1. Прочитать все значения из диапазона A2:A7.
2. Найти значения, которые встречаются в этом списке более 1 раза (в данном примере это `101` и `103`).
3. Для **всех** ячеек, в которых встречается найденный дубликат, установить цвет заливки `#FF0000` (Красный).
4. Уникальные ID (в данном примере `102` и `104`) должны остаться без изменения цвета.

*Подсказка: Для создания объекта цвета используйте `Api.CreateColorFromHex("#FF0000")`.*
