## Обзор

Используйте Telegram для обмена мгновенными сообщениями в интерфейсе редакторов R7-Office. 

Плагин основан на приложении [telegram-react](https://github.com/evgeny-nadymov/telegram-react). Приложение использует среду JavaScript ReactJS и TDLib (библиотеку базы данных Telegram), скомпилированную в WebAssembly. 

Плагин совместим с [самостоятельными](https://github.com/R7-Office/DocumentServer) и [настольными](https://github.com/R7-Office/DesktopEditors) версиями редакторов R7-Office. Его можно добавить в экземпляры R7-Office вручную. 

## Как использовать

1. Найдите плагин на вкладке Плагины.
2. Войдите в свою учетную запись Telegram. 

## Как установить

Подробные инструкции можно найти в [Документации по R7-Office API](https://api.R7-Office.com/docs/plugin-and-macros/tutorials/installing/R7-Office-docs-on-premises/).

## Конфигурация

По умолчанию этот плагин использует (https://evgeny-nadymov.github.io/telegram-react/). Если вам нужно изменить его, откройте файл `index.html` и вставьте новый URL-адрес в iframe `src field`.

## Известные проблемы

* Плагин не имеет доступа к камере и микрофону, поэтому вы не сможете записывать голосовые и видеосообщения. 
* Плагин не работает в режиме инкогнито. 


---
создано при поддержке [https://r7-consult.ru/](https://r7-consult.ru/)