## Overview

Use Telegram for instant messaging within R7-Office editors interface. 

The plugin is based on the [telegram-react](https://github.com/evgeny-nadymov/telegram-react) app. The app uses the ReactJS JavaScript framework and TDLib (Telegram Database library) compiled to WebAssembly. 

The plugin is compatible with [self-hosted](https://github.com/R7-Office/DocumentServer) and [desktop](https://github.com/R7-Office/DesktopEditors) versions of R7-Office editors. It can be added to R7-Office instances manually. 

## How to use

1. Find the plugin in the Plugins tab.
2. Log in to your Telegram account. 

## How to install

Detailed instructions can be found in [R7-Office API documentation](https://api.R7-Office.com/docs/plugin-and-macros/tutorials/installing/R7-Office-docs-on-premises/).

## Configuration

By default, that plugin uses (https://evgeny-nadymov.github.io/telegram-react/). If you need to change it, open the `index.html` file and insert the new URL in the iframe `src field`.

## Known issues

* The plugin has no access to the camera and microphone, so you'll be unable to record voice and video messages. 
* The plugin doesn't work in the incognito mode. 


---
создано при поддержке [https://r7-consult.ru/](https://r7-consult.ru/)