<p align="center">
 <img width="467px" height="115" alt="js-excel-generator" src="https://i.imgur.com/Xw7Xfnv.png"/> 
</p>
<p align="center">
<b>A JavaScript Library for creating Excel Spreadsheets from HTML Tables</b>
</p>

## Overview
A Small HTML JavaScript Library which can be used to create Excel Spreadsheets 
from HTML Tables.  

[Live Demo](https://rawgit.com/ecscstatsconsulting/js-excel-generator/master/demo.htm)

## Getting Started - How to use

 - Include and reference the excel-gen.js library and the jquery, jszip and FileSaver dependencies
 
    example:
 ```
    <script type="text/javascript" src="./external/jquery-1.8.2.js"></script>
    <script type="text/javascript" src="./external/jszip.js"></script>
    <script type="text/javascript" src="./external/FileSaver.js"></script>
    <script type="text/javascript" src="./scripts/excel-gen.js"></script>
 ```
 - Call the library within on the desired HTML Table
    
    example:
    
 ```javascript
 $(document).ready(function () {
    //Creates new Generator
    excel = new ExcelGen({
        "src_id": "[ID For Table]",
        "show_header": true,
        "file_name": "myExcelOutput.xlsx"
    });
    //Generates Excel Output and sends download to the browser.
    excel.generate();
});
 ```
 
## Usage - Options
The following options are current available:
```
{
    "src_id": "",
    "header_row": null,
    "body_rows": null,
    "type": "table|normal",
    "show_header": false,
    "author": "JavaScript Excel Generator",
    "file_name": "output.xlsx",
    "column_formats": []
}
```

**src_id** - ID of HTML Table to Export to Excel.  Can use header_row and body_rows instead.

**header_row** - JQuery Element of the header TR tag, if src_id is not specified this will be used.

**body_rows** - Jquery Elements of the body TR tags, if src_id is not specified this will be used.

**__note: both header_row and body_rows properties must be set if not using src_id__**

**type** - If set to 'table', the output will be generated inside of a Filterable Excel Table

**show_header** - Show the header on the Excel Output

**author** - Name to be listed as the author of the Excel Output

**file_name** - File name of the Excel Output

**column_formats** - Array of column formats (see [demo](https://rawgit.com/ecscstatsconsulting/js-excel-generator/master/demo.htm) for more info).  For a list of column formats see [ExcelDataFormats.md](ExcelDataFormats.md)

## Dependencies
The code uses jQuery, jszip.js and FileSaver.js.  For ease of use, these libraries 
have been included in the 'external' folder of this project.

<p align="center">
 <img width="401px" height="106" alt="ecsc" src="https://i.imgur.com/SzVdycv.png"/> 
</p>
