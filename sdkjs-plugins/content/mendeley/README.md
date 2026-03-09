# R7-Office Mendeley plugin

Mendeley plugin allows users to create bibliographies in R7-Office editors using [Mendeley service](https://www.mendeley.com/).

The plugin is pre-installed in [R7-Office Workspace](https://www.R7-Office.com/workspace.aspx) (both Enterprise and Community Edition), [R7-Office cloud service](https://www.R7-Office.com/cloud-office.aspx), and [R7-Office Personal](https://personal.R7-Office.com/). It can also be installed to [Document Server](https://github.com/R7-Office/DocumentServer) manually.

## How to use

1. Search references by author, title or year.

2. Among search results, choose ones you want to add to your document.

3. Choose style (e.g. Chicago Manual, American Psychological Association) and language.

4. Press `Insert citation`.

## How to install

Two installation ways are available:

1. Put the folder with Mendeley plugin (it must contain the content of the src folder only) to R7-Office Document Server folder depending on the operating system:

    For Linux - `/var/www/R7-Office/documentserver/sdkjs-plugins/`.

    For Windows - `%ProgramFiles%\R7-Office\DocumentServer\sdkjs-plugins\`.

    The plugins will be available to all the users users of R7-Office Document Server.
    No service restart is required.

2. Edit the Document Server config to add the following lines:

    ```
    var docEditor = new DocsAPI.DocEditor("placeholder", {
        "editorConfig": {
            "plugins": {
                "autostart": [
                    "asc.{BE5CBF95-C0AD-4842-B157-AC40FEDD9441}",
                    ...
                ],
                "pluginsData": [
                    "https://example.com/path/to/mendeley/config.json",
                    ...
                ]
            },
            ...
        },
        ...
    });
    ```

Detailed instructions can also be found in [R7-Office API documentation](https://api.R7-Office.com/docs/plugin-and-macros/tutorials/installing/R7-Office-docs-on-premises/).

**Important**: when you integrate R7-Office Document Server with a 3rd-party storage, you need to use [special connectors](https://api.R7-Office.com/editors/plugins) (integration apps). If you compile a connector from source code or create a new one, you can add plugins using Document Server config. If you use ready connectors (e.g. from ownCloud/Nextcloud marketplaces) adding plugins via config is not applicable. 

## Configuration

You will need to register the application.

1. Go to https://dev.mendeley.com/myapps.html.

2. Fill in the form using link provided in the plugin interface as a redirect URL.

3. Press `Generate secret` and copy it.

4. Insert the secret into the appropriate field in the plugin interface.

## Known issues

For CentOS users with SELinx enabled, after copying the src folder to sdkjs-plugins, plugins may not work due to the variable file security context. To restore the rights, use the following command:

```
sudo restorecon -Rv /var/www/R7-Office/documentserver/sdkjs-plugins/
```

After that restart the services:

```
sudo supervisorctl restart ds:docservice
```

This plugin doesn't work in desktop editor, because has problem with authorization. We know about this problem and will try to fix it in the feature.

## User feedback and support

To ask questions and share feedback, use Issues in this repository.

If you need more information about how to use or write your own plugin, please visit our [API documentation](https://api.R7-Office.com/docs/plugin-and-macros/get-started/overview/).

---
создано при поддержке [https://r7-consult.ru/](https://r7-consult.ru/)