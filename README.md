# ts-xlsx-export


This library allows you to export json data in Excel. It has been written in typescripte so it can be used in all projects using typescript. For example, projects with Angular. The installation and use are very simple.. 

## Installation

```Javascript
    npm i ts-xlsx-export
```

## Utilization

```Javascript
    import { tsXLXS } from 'ts-xlsx-export';

    tsXLXS().exportAsExcelFile(jsonData).saveAsExcelFile(fileName);
```

## Function details

<table style="text-align: left">
<tr>
    <th>Function</th>
    <th>Desciption</th>
    <th>Params</th>
</tr>
<tr>
    <td>tsXLXS()</td>
    <td>the first function called that allow to instance the library</td>
    <td> Null</td>
</tr>
<tr>
    <td>exportAsExcelFile()</td>
    <td>this function prepares json data to be downloaded in excel format</td>
    <td>Object Array</td>
</tr>
<tr>
    <td>saveAsExcelFile()</td>
    <td>function that allow to download and save file, your can pass the param for the name of the file</td>
    <td>String</td>
</tr>
</table>


## Browser Support

This library has been tested with the following browsers:
* Chrome
* IE 11
* FireFox

