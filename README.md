# Service for importing data from external ONLYOFFICE spreadsheets

[ONLYOFFICE](https://www.onlyoffice.com/) doesn't have the possiblity of importing data from external spreadsheets (e.g. ImportRange in Google). This small web-service works together with a [plugin](https://github.com/dandm1986/Import_Data_Plugin_ONLYOFFICE) installed in a desktop ONLYOFFICE version.

`Endpoint` https://import-data-onlyoffice.herokuapp.com/api/v1/methods/import-range

## Request structure

`Get an access token`

```

```

`POST api/v1/methods/import-range`

```
{
    "token": "hv89dvamkuKiD/0VCGkLzJAPCPYg8XLAiYv8sDmaei8xKrdK5u3ZofV1oyHy+G3NOgMqWYNbRIIE8LIAQlSke0AXjNMlRxx16cfKQoa60W6RiMMoJroSEQjhAgUTP3adXqEN0cMm+UsRt1pPYns9u+Nby1qpAFg8Y9mz8a1lMs=",
    "urlToFile": "url_to_onlyoffice_server/Products/Files/DocEditor.aspx?fileid=8120993&doc=N0UwWmdvdElNQmlZWnl3cXBtQVN0cmlHWkFZQk1PTVJ1R2VSTUxDcUY3cz0_IjgxMjA5OTMi0",
    "worksheet": "Sheet1",
    "range": "A1:B10",
    "targetCell": "C1"
}
```

`Response example`

```
{
  "status": "success",
  "data": {
      "cellsObj": {
          "C1": {
              "value": "General",
              "format": "General"
          },
          "D1": {
              "value": 1000,
              "format": "General"
          },
          "C2": {
              "value": "Number",
              "format": "General"
          },
          "D2": {
              "value": 1000,
              "format": "0.00"
          },
          "C3": {
              "value": "Scientific",
              "format": "General"
          },
          "D3": {
              "value": 1000,
              "format": "0.00E+00"
          },
          "C4": {
              "value": "Accounting",
              "format": "General"
          },
          "D4": {
              "value": 1000,
              "format": "_($* #,##0.00_)"
          },
          "C5": {
              "value": "Currency",
              "format": "General"
          },
          "D5": {
              "value": 1000,
              "format": "$#,##0.00"
          },
          "C6": {
              "value": "Date",
              "format": "General"
          },
          "D6": {
              "value": 1000,
              "format": "m/d/yy"
          },
          "C7": {
              "value": "Time",
              "format": "General"
          },
          "D7": {
              "value": 1000,
              "format": "[$-F400]h:mm:ss AM/PM"
          },
          "C8": {
              "value": "Percentage",
              "format": "General"
          },
          "D8": {
              "value": 1000,
              "format": "0.00%"
          },
          "C9": {
              "value": "Fraction",
              "format": "# ?/?"
          },
          "D9": {
              "value": 1000,
              "format": "# ?/?"
          },
          "C10": {
              "value": "Text",
              "format": "General"
          },
          "D10": {
              "value": 1000,
              "format": "@"
          }
      }
  }
}
```
