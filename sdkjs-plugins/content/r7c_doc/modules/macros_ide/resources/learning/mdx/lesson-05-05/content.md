# Урок 5.5: Временная декомпозиция и сезонный анализ

Введение

Представьте владельца магазина игрушек, который в декабре 2023 года продал товаров на 40% больше, чем в ноябре. Успех? Не спешите с выводами. Рождественские продажи традиционно составляют до 40% годовой выручки в этом бизнесе. Возможно, рост в 40% — это провал, если обычно декабрь показывает +60%. Как отделить реальный рост бизнеса от сезонного пика? Как понять, действительно ли декабрь был успешным? Ответ — в декомпозиции временного ряда, разложении данных на составляющие компоненты.

Сезонность окружает нас повсюду. Туристические компании видят пики летом и провалы зимой. Энергетики фиксируют всплески потребления электричества летом из-за кондиционеров и зимой из-за отопления. Ритейлеры живут от "черной пятницы" до новогодних распродаж. Даже рестораны быстрого питания знают, что пятница вечером принесет в три раза больше заказов, чем понедельник утром. Без понимания этих паттернов невозможно отличить естественные колебания от реальных проблем или успехов.

Декомпозиция временного ряда — это как снятие нескольких слоев с луковицы. Мы отделяем долгосрочный тренд (растет ли бизнес в целом?) от сезонности (ожидаемые колебания) и случайных факторов (разовые акции, погодные аномалии). Это позволяет принимать правильные решения: не паниковать при сезонном январском спаде после новогодних праздников и не переоценивать декабрьский подъем. В этом финальном уроке модуля мы объединим все изученные временные функции MDX для профессионального анализа сезонности.

Теоретическая часть

A. Компоненты временного ряда

## Любой временной ряд можно представить как комбинацию четырех компонентов

Тренд (Trend) — долгосрочное направление развития, очищенное от всех колебаний. Это ответ на вопрос "Куда движется бизнес?"

Сезонность (Seasonal) — регулярные, предсказуемые колебания с фиксированным периодом (год, квартал, месяц, неделя).

Циклическая компонента (Cyclical) — долгосрочные волны без фиксированного периода, связанные с бизнес-циклами экономики.

Случайная компонента (Irregular) — непредсказуемые флуктуации: разовые акции, погодные аномалии, форс-мажоры.

## Модели комбинирования

Аддитивная: Y = Trend + Seasonal + Irregular (когда амплитуда сезонности постоянна)

Мультипликативная: Y = Trend × Seasonal × Irregular (когда амплитуда растет с уровнем)

💡 Pro Tip: Выбор между аддитивной и мультипликативной моделью

Если сезонные колебания постоянны по амплитуде → аддитивная

Если амплитуда растет с ростом уровня → мультипликативная Простой тест: постройте график, если "конверт" расширяется — используйте мультипликативную модель.

B. Методы выявления сезонности в MDX

## MDX предоставляет богатый инструментарий для анализа сезонности

ParallelPeriod — сравнение с аналогичным периодом прошлого года устраняет сезонность

Расчет сезонных индексов — отношение среднего за конкретный месяц к общему среднему

LastPeriods + скользящие средние — выделение тренда через сглаживание

Year-over-Year анализ — фокус на изменениях, а не абсолютных значениях

C. Расчет сезонных коэффициентов

## Классический метод

Рассчитать 12-месячную скользящую среднюю (устраняет сезонность)

Найти отношение фактических значений к скользящей средней

Усреднить отношения для каждого месяца за несколько лет

Нормализовать, чтобы сумма индексов = 12

D. Десезонализация данных

Десезонализированные данные показывают, какими были бы продажи без сезонного эффекта. Это критически важно для:

Оценки реального роста бизнеса

Сравнения разносезонных периодов

Выявления трендов и точек перелома

Обнаружения аномалий

E. Прогнозирование с учетом сезонности

Качественный прогноз = Экстраполяция тренда × Сезонный коэффициент × Поправка на особые факторы

F. Интеграция с предыдущими уроками

## Каждая изученная функция играет свою роль в декомпозиции

Lag/Lead (5.1) — расчет приростов и темпов изменения

ParallelPeriod (5.2) — year-over-year анализ без сезонности

PeriodsToDate (5.3) — накопительные итоги для оценки годовых трендов

LastPeriods (5.4) — скользящие средние для выделения тренда

## Визуализация декомпозиции

AVR Assembly

Original: ╱╲╱╲╱╲╱╲  (тренд + сезонность + шум)

Trend:    ╱╱╱╱╱╱╱╱   (выделенный тренд)

Seasonal: ∩∪∩∪∩∪∩∪   (сезонный паттерн)

Random:   -~-~-~-~   (остаток)

Практическая часть

Пример 1: Расчет сезонных индексов

```mdx
WITH
-- Среднее значение для каждого месяца за 3 года (2011-2013)
MEMBER [Measures].[Jan Avg] AS
    AVG({
        ([Date].[Calendar].[Calendar Year].&[2011].FirstChild.FirstChild, [Measures].[Internet Sales Amount]),
        ([Date].[Calendar].[Calendar Year].&[2012].FirstChild.FirstChild, [Measures].[Internet Sales Amount]),
        ([Date].[Calendar].[Calendar Year].&[2013].FirstChild.FirstChild, [Measures].[Internet Sales Amount])
    })
-- Аналогично для других месяцев (показываем несколько для примера)
MEMBER [Measures].[Jun Avg] AS
    AVG({
        ([Date].[Calendar].[Calendar Year].&[2011].Children.Item(1).LastChild, [Measures].[Internet Sales Amount]),
        ([Date].[Calendar].[Calendar Year].&[2012].Children.Item(1).LastChild, [Measures].[Internet Sales Amount]),
        ([Date].[Calendar].[Calendar Year].&[2013].Children.Item(1).LastChild, [Measures].[Internet Sales Amount])
    })
MEMBER [Measures].[Dec Avg] AS
    AVG({
        ([Date].[Calendar].[Calendar Year].&[2011].LastChild.LastChild, [Measures].[Internet Sales Amount]),
        ([Date].[Calendar].[Calendar Year].&[2012].LastChild.LastChild, [Measures].[Internet Sales Amount]),
        ([Date].[Calendar].[Calendar Year].&[2013].LastChild.LastChild, [Measures].[Internet Sales Amount])
    })
-- Общее среднее за все месяцы
MEMBER [Measures].[Overall Monthly Avg] AS
    AVG(
        DESCENDANTS(
            {[Date].[Calendar].[Calendar Year].&[2011],
             [Date].[Calendar].[Calendar Year].&[2012],
             [Date].[Calendar].[Calendar Year].&[2013]},
            [Date].[Calendar].[Month],
            SELF
        ),
        [Measures].[Internet Sales Amount]
    )
-- Сезонный индекс для текущего месяца
MEMBER [Measures].[Seasonal Index] AS
    CASE [Date].[Calendar].CurrentMember.Name
        WHEN "January" THEN [Measures].[Jan Avg] / [Measures].[Overall Monthly Avg]
        WHEN "June" THEN [Measures].[Jun Avg] / [Measures].[Overall Monthly Avg]
        WHEN "December" THEN [Measures].[Dec Avg] / [Measures].[Overall Monthly Avg]
        -- Добавить остальные месяцы
        ELSE 1.0
    END,
    FORMAT_STRING = "#,##0.00"
-- Интерпретация индекса
MEMBER [Measures].[Index Interpretation] AS
    CASE
        WHEN [Measures].[Seasonal Index] > 1.2 THEN "High Season (+" +
            STR(ROUND(([Measures].[Seasonal Index] - 1) * 100, 0)) + "%)"
        WHEN [Measures].[Seasonal Index] < 0.8 THEN "Low Season (" +
            STR(ROUND(([Measures].[Seasonal Index] - 1) * 100, 0)) + "%)"
```

        ELSE "Normal Season"

```mdx
    END
SELECT
    {[Measures].[Internet Sales Amount],
     [Measures].[Seasonal Index],
     [Measures].[Index Interpretation]} ON COLUMNS,
    DESCENDANTS(
        [Date].[Calendar].[Calendar Year].&[2013],
        [Date].[Calendar].[Month],
        SELF
    ) ON ROWS
FROM [Adventure Works]
```

Пример 2: Декомпозиция временного ряда

```mdx
WITH
-- Тренд: 12-месячная центрированная скользящая средняя
MEMBER [Measures].[Trend Component] AS
    IIF(
        COUNT(LastPeriods(12, [Date].[Calendar].CurrentMember), EXCLUDEEMPTY) < 12,
        NULL,
        AVG(
            LastPeriods(12, [Date].[Calendar].CurrentMember),
            [Measures].[Internet Sales Amount]
        )
    ),
    FORMAT_STRING = "Currency"
-- Детрендированные данные (факт / тренд для мультипликативной модели)
MEMBER [Measures].[Detrended] AS
    IIF(
        ISEMPTY([Measures].[Trend Component]) OR [Measures].[Trend Component] = 0,
        NULL,
        [Measures].[Internet Sales Amount] / [Measures].[Trend Component]
    ),
    FORMAT_STRING = "#,##0.00"
```

-- Сезонная компонента (упрощенно - как среднее отклонение от тренда)

```mdx
MEMBER [Measures].[Seasonal Component] AS
    -- Для полной реализации нужно усреднить Detrended по месяцам за несколько лет
    [Measures].[Detrended],
    FORMAT_STRING = "#,##0.00"
-- Случайная компонента (остаток после удаления тренда и сезонности)
MEMBER [Measures].[Irregular Component] AS
    IIF(
        ISEMPTY([Measures].[Trend Component]),
        NULL,
        [Measures].[Internet Sales Amount] -
        ([Measures].[Trend Component] * [Measures].[Seasonal Component])
    ),
    FORMAT_STRING = "Currency"
-- Качество декомпозиции (доля объясненной вариации)
MEMBER [Measures].[R-Squared] AS
    1 - (
        [Measures].[Irregular Component] * [Measures].[Irregular Component] /
        (([Measures].[Internet Sales Amount] - [Measures].[Trend Component]) *
         ([Measures].[Internet Sales Amount] - [Measures].[Trend Component]))
    ),
    FORMAT_STRING = "Percent"
SELECT
    {[Measures].[Internet Sales Amount],
     [Measures].[Trend Component],
     [Measures].[Seasonal Component],
     [Measures].[Irregular Component],
     [Measures].[R-Squared]} ON COLUMNS,
    DESCENDANTS(
        [Date].[Calendar].[Calendar Year].&[2013],
        [Date].[Calendar].[Month],
        SELF
    ) ON ROWS
FROM [Adventure Works]
```

Пример 3: Десезонализация и анализ реального роста

```mdx
WITH
-- Базовый сезонный индекс (упрощенный расчет для демонстрации)
MEMBER [Measures].[Monthly Seasonal Index] AS
    CASE MONTH([Date].[Calendar].CurrentMember.Name)
```

        WHEN 1 THEN 0.85  -- Январь обычно -15%

        WHEN 6 THEN 1.05  -- Июнь обычно +5%

        WHEN 11 THEN 1.25 -- Ноябрь обычно +25%

        WHEN 12 THEN 1.40 -- Декабрь обычно +40%

        ELSE 1.0

    END,

```mdx
    FORMAT_STRING = "#,##0.00"
-- Десезонализированные продажи
MEMBER [Measures].[Deseasonalized Sales] AS
    [Measures].[Internet Sales Amount] / [Measures].[Monthly Seasonal Index],
    FORMAT_STRING = "Currency"
-- Реальный рост (месяц к месяцу, десезонализированный)
MEMBER [Measures].[Real MoM Growth] AS
    IIF(
        ISEMPTY(([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Deseasonalized Sales])),
        NULL,
        ([Measures].[Deseasonalized Sales] -
         ([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Deseasonalized Sales])) /
        ([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Deseasonalized Sales])
    ),
    FORMAT_STRING = "Percent"
-- Номинальный рост для сравнения
MEMBER [Measures].[Nominal MoM Growth] AS
    IIF(
        ISEMPTY(([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Internet Sales Amount])),
        NULL,
        ([Measures].[Internet Sales Amount] -
         ([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Internet Sales Amount])) /
        ([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Internet Sales Amount])
    ),
    FORMAT_STRING = "Percent"
-- Детекция аномалий (отклонение от ожидаемого сезонного уровня)
MEMBER [Measures].[Anomaly Score] AS
    ABS([Measures].[Internet Sales Amount] -
        ([Measures].[Deseasonalized Sales] * [Measures].[Monthly Seasonal Index])) /
    [Measures].[Internet Sales Amount],
    FORMAT_STRING = "Percent"
MEMBER [Measures].[Anomaly Status] AS
    CASE
        WHEN [Measures].[Anomaly Score] > 0.2 THEN "⚠️ Major Anomaly"
        WHEN [Measures].[Anomaly Score] > 0.1 THEN "Minor Deviation"
        ELSE "Normal"
    END
SELECT
    {[Measures].[Internet Sales Amount],
     [Measures].[Monthly Seasonal Index],
     [Measures].[Deseasonalized Sales],
     [Measures].[Nominal MoM Growth],
     [Measures].[Real MoM Growth],
     [Measures].[Anomaly Score],
     [Measures].[Anomaly Status]} ON COLUMNS,
    DESCENDANTS(
        [Date].[Calendar].[Calendar Year].&[2013],
        [Date].[Calendar].[Month],
        SELF
    ) ON ROWS
FROM [Adventure Works]
```

Пример 4: Комплексный прогноз с учетом тренда и сезонности

```mdx
WITH
```

-- Линейный тренд (упрощенный - прирост за последние 3 месяца)

```mdx
MEMBER [Measures].[Trend Growth Rate] AS
    AVG(
        LastPeriods(3, [Date].[Calendar].CurrentMember),
        ([Measures].[Internet Sales Amount] -
         ([Date].[Calendar].CurrentMember.Lag(1), [Measures].[Internet Sales Amount]))
    ),
    VISIBLE = 0
-- Прогноз тренда на следующий месяц
MEMBER [Measures].[Trend Forecast] AS
    [Measures].[Internet Sales Amount] + [Measures].[Trend Growth Rate],
    FORMAT_STRING = "Currency"
-- Сезонный множитель следующего месяца
MEMBER [Measures].[Next Month Season] AS
    CASE MONTH([Date].[Calendar].CurrentMember.Lead(1).Name)
```

        WHEN 1 THEN 0.85

        WHEN 12 THEN 1.40

        ELSE 1.0

    END,

```mdx
    FORMAT_STRING = "#,##0.00"
-- Финальный прогноз = тренд × сезонность
MEMBER [Measures].[Final Forecast] AS
    [Measures].[Trend Forecast] * [Measures].[Next Month Season],
    FORMAT_STRING = "Currency"
-- Доверительный интервал (±15% для примера)
MEMBER [Measures].[Forecast Lower Bound] AS
    [Measures].[Final Forecast] * 0.85,
    FORMAT_STRING = "Currency"
MEMBER [Measures].[Forecast Upper Bound] AS
    [Measures].[Final Forecast] * 1.15,
    FORMAT_STRING = "Currency"
-- Оценка точности (если есть фактические данные)
MEMBER [Measures].[Forecast Accuracy] AS
    IIF(
        ISEMPTY(([Date].[Calendar].CurrentMember.Lead(1), [Measures].[Internet Sales Amount])),
        "No Actual Data",
        STR(ROUND(
            (1 - ABS(
                [Measures].[Final Forecast] -
                ([Date].[Calendar].CurrentMember.Lead(1), [Measures].[Internet Sales Amount])
            ) / ([Date].[Calendar].CurrentMember.Lead(1), [Measures].[Internet Sales Amount])
```

        ) * 100, 1)) + "% Accurate"

```mdx
    )
SELECT
    {[Measures].[Internet Sales Amount],
     [Measures].[Trend Forecast],
     [Measures].[Next Month Season],
     [Measures].[Final Forecast],
     [Measures].[Forecast Lower Bound],
     [Measures].[Forecast Upper Bound],
     [Measures].[Forecast Accuracy]} ON COLUMNS,
    HEAD(
        DESCENDANTS(
            [Date].[Calendar].[Calendar Year].&[2013],
            [Date].[Calendar].[Month],
            SELF
        ),
        6
    ) ON ROWS
FROM [Adventure Works]
```

⚠️ Важно: Сезонность может меняться со временем! Пандемия COVID-19 полностью изменила сезонные паттерны во многих индустриях. Всегда проверяйте актуальность исторических сезонных коэффициентов.

## Таблица сезонных индексов (пример)

Месяц

Индекс

Интерпретация

Январь

0.85

```mdx
-15% от среднего
```

Февраль

0.90

```mdx
-10% от среднего
```

Июнь

1.05

```mdx
+5% от среднего
```

Ноябрь

1.25

```mdx
+25% от среднего
```

Декабрь

1.40

```mdx
+40% от среднего
```

## График сезонности

Coq

## Sales Index by Month

1.4 |                      *Dec

1.2 |                  *Nov

1.0 |---*Jun---*Oct-----------

0.8 | *Jan  *Feb

Заключение

Временная декомпозиция — это кульминация всех техник временного анализа, которые мы изучили в этом модуле. Она позволяет видеть истинную картину бизнеса, отделяя устойчивый рост от временных всплесков, ожидаемую сезонность от неожиданных аномалий. Правильное понимание сезонности критически важно для планирования запасов, бюджетирования маркетинга и оценки эффективности менеджмента.

Мы научились использовать весь арсенал временных функций MDX: ParallelPeriod для устранения сезонности через year-over-year сравнения, PeriodsToDate для накопительного анализа трендов, LastPeriods для сглаживания и выделения долгосрочных паттернов. Каждая функция вносит свой вклад в комплексное понимание временных рядов.

В следующем модуле мы углубимся в работу с иерархиями и навигацию по многомерным структурам, что позволит применять изученные временные техники к любым измерениям бизнес-данных.
