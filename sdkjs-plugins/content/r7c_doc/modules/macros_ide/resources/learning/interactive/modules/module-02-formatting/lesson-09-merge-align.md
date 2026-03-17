---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-format-09"
  title: "Урок 9: Объединение (Merge) и Alignment"
  draft: false
  beforeScript: |
    (function() {
    let sheet = Api.GetActiveSheet();
    sheet.GetRange("A1").SetValue("Главный Заголовок Отчета");
    sheet.GetRange("A1:H1").SetFillColor(Api.CreateColorFromRGB(220, 230, 241));
    sheet.GetRange("A4:H4").Merge(false);
    sheet.GetRange("A4:H4").SetValue("💡 Напишите макрос, чтобы объединить A1:H1 и выровнять текст по центру.");
    })();
---

# Урок 9. Программная верстка: Range.Merge и Alignment Properties

Визуальное форматирование (раскраска) ячеек полезно для дашбордов. Однако, при автоматизированной генерации отчетности (Dynamic Reporting Templates) из сырых (RAW) данных вам часто необходимо программно управлять макетом таблицы: объединять заголовочные ячейки, формировать рамки и подписи, выравнивать контент.

## 🎯 Слияние Ячеек: Метод `Merge`
В VBA консолидация блоков выполнялась через `Range.MergeCells = True` или `Range.Merge`.
В API Р7-Офис используется явно типизированный метод `.Merge(boolean)` вызванный у инстанса `ApiRange`.

Аргумент метода определяет паттерн слияния:
- `.Merge(false)` — классическое блочное слияние всех ячеек выделенного Range в одну гигантскую ячейку. *(Аналог VBA: `Merge(Across:=False)`)*.
- `.Merge(true)` — построчное слияние. Объединяет каждую выделенную строку отдельно.

### Синтаксический пример
```javascript
let sheet = Api.GetActiveSheet();
let headerRange = sheet.GetRange("A1:E1"); // Декларативный массив ячеек

// Инстанцирование слияния в прямоугольник
headerRange.Merge(false); 
```

---

## 📐 Выравнивание текста (Alignment API)

В макросах VBA выравнивание `Alignment` управляется константами наподобие `xlCenter`, `xlRight`, `xlTop`.
Движок Р7-Офис, работающий по принципам Web-стандартов, принимает строковые аргументы (String Literals), аналогичные CSS-свойствам (`justify`, `center`).

Для модификации осей в Р7-Офис предназначены методы (Сеттеры):
- Перехват оси X: `.SetAlignHorizontal(horizontalAlign)`
- Перехват оси Y: `.SetAlignVertical(verticalAlign)`

| Ось | Строковые Константы P7 (JS) | Аналоги Констант VBA (MS Excel) |
| --- | :--- | :--- |
| **Ось X** | `"center"` / `"left"` / `"right"` | `xlCenter` / `xlLeft` / `xlRight` |
| **Ось Y** | `"center"` / `"top"` / `"bottom"` | `xlCenter` / `xlTop` / `xlBottom`|

### 📜 Архитектура генератора шапки таблицы
```javascript
(function() {
    'use strict';
    try {
        const api = Api;
        let sheet = Api.GetActiveSheet();
        let header = sheet.GetRange("A1:E1");
        
        header.Merge(false);
        
        // Инициализация значения ПОСЛЕ слияния
        header.SetValue("ВНУТРЕННИЙ АУДИТ 2026");
        
        header.SetBold(true);
        header.SetFontSize(20);
        
        // UI Центрирование
        header.SetAlignHorizontal("center"); 
        header.SetAlignVertical("center");
    } catch(e) {}
})();
```

---

## ⚠️ Ловушки и лучшие практики

1. **Последовательность операций (Data Loss при Merge)**
   Как и в Microsoft Excel, слияние в Р7-Офис уничтожает все значения (Values) ячеек, кроме верхней левой. **Лучшая практика:** Сначала делайте `header.Merge(false)`, и ТОЛЬКО ЗАТЕМ вписывайте значение (String/Number) через `.SetValue()`, чтобы не потерять расчетные данные.

2. **Запрет на объединение базы данных (DB Constraint)**
   *Никогда не выполняйте `Merge` внутри строк с сырыми данными (Rows data).* 
   Объединенные ячейки необратимо разрушают алгоритмы сортировок, фильтров и авто-построения Сводных Таблиц (Pivot Tables). Делайте слияние исключительно в заголовочных конструкциях (Headers/Footers).

Изучив методы `Merge` и `Align`, переходите к следующей практике (Урок 10), где вы сами сгенерируете Enterprise-шапку для отчета.
