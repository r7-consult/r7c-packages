# Slider Query

`Slider Query` is a commercial self-service analytics plugin for the R7 Spreadsheet Editor. It is designed for everyday reporting, SQL-based data preparation, and OLAP analysis directly inside R7-Office, without switching to external BI tools for routine work.

The product is developed by `Datacons` and positioned as a practical working alternative for scenarios that are often covered by Power Query, Microsoft Access, or Power BI in other office environments.

## What the plugin is for

`Slider Query` helps analysts and business users work with operational and analytical data in a familiar spreadsheet interface. Instead of exporting files manually, combining data in several tools, and rebuilding reports from scratch, users can prepare and refresh reports inside one workflow.

The plugin is focused on:

- regular reporting and operational analytics;
- combining data from multiple systems in one report;
- building reusable SQL-based report logic;
- interactive work with OLAP cubes and pivot-style analytical views.

## Core capabilities

### Connection Manager

The plugin includes a dedicated connection manager for configuring and validating data sources. It supports work with relational databases, local files, and OLAP connections.

Typical source types include:

- PostgreSQL, MS SQL, MySQL and other relational databases;
- Excel, CSV/TXT, JSON, Parquet and file directories;
- OLAP/XMLA sources for multidimensional analytics.

It also supports import and export of connection settings, which is useful for transfer between environments and controlled setup of working configurations.

### SQL Manager

The SQL manager is the main workspace for building and maintaining report logic. It allows users to create queries, organize them by sources, preview results, and reuse prepared logic inside regular workflows.

The SQL manager supports several practical capabilities:

- creation of standard SQL reports against configured sources;
- preview of query results before loading data to a sheet;
- use of variables for dynamic reports;
- support for heterogeneous queries that combine data from different systems;
- building chained transformations where one prepared dataset becomes the source for the next step.

### Heterogeneous data workflows

One of the strongest practical features is the ability to combine data from different source types in a single analytical flow. In real work this is important when a report depends on data that lives in different systems and formats.

This allows teams to:

- join database tables with local files;
- prepare consolidated datasets for regular reporting;
- move part of the processing closer to the source systems and load only the final result into the editor.

### OLAP and pivot-style analytics

For multidimensional analysis, the plugin provides a separate OLAP module. It supports interactive work with cubes, field selection, filters, and report layouts directly from the spreadsheet environment.

This is relevant for:

- ad-hoc analytical exploration;
- management reporting based on OLAP cubes;
- quick slicing of measures and dimensions without rebuilding the report manually.

## Typical use cases

`Slider Query` is especially useful when:

- reports have to be refreshed regularly from the same sources;
- data is distributed across several systems and formats;
- users need more control than standard spreadsheet imports provide;
- analysts want BI-style capabilities while staying inside R7 Spreadsheet Editor.

## Support

- Website: `https://data.slider-ai.ru/`
- Email: `data@slider-ai.ru`
- Telegram: `https://t.me/SliderQuery`
