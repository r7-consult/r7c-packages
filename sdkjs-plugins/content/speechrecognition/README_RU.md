## Обзор

Введите голос в R7-Office Docs.

Плагин использует [Web Speech API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Speech_API).

Плагин можно установить вручную в [собственно размещенную](https://github.com/R7-Office/DocumentServer) версию R7-Office Docs.

## Как использовать

1. Включите микрофон.
2. Найдите плагин на вкладке Плагины и щелкните его.
3. Установите язык, который вы используете.
4. Нажмите кнопку микрофона и начните говорить. 
5. После завершения диктанта текст вставляется в документ. Чтобы прекратить вставку текста, нажмите кнопку микрофона еще раз. 

## Как установить

Подробные инструкции можно найти в [Документации по R7-Office API](https://api.R7-Office.com/docs/plugin-and-macros/tutorials/installing/R7-Office-docs-on-premises/).

## Известные проблемы

- Этот плагин работает только в Google Chrome. Чтобы использовать его в FireFox, включите распознавание с помощью флагов `media.webspeech.recognition.enable` и `media.webspeech.recognition.force_enable`, в `about:config` синтез включен по умолчанию. Дополнительная информация [здесь](https://developer.mozilla.org/en-US/docs/Web/API/Web_Speech_API). 

- Плагин пока не работает для десктопной версии.

- Плагин не работает с http (работает только с https).



---
создано при поддержке [https://r7-consult.ru/](https://r7-consult.ru/)