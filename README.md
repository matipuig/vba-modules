# VBA Modules

This repo contains some modules we use often at work for offices files with macros.
There are some functions in VBA which are difficult or verbose to use (like regex, searching text, etc.).
The aim of this repo is to make all these functions easier to use.

There's also a file which use this modules to complete doc Word templates from an Excel table in useful files.

## Usage

Just press "Import" in the VBA IDE and add this file. It will be imported with the same name.

## Modules.

| Module          | Description                                                                  |
| --------------- | ---------------------------------------------------------------------------- |
| mdlAdobeAcrobat | It works if you have Adobe Acrobat. It can open it and read files.           |
| mdlArrays       | It lets you add to array, join and search in them (Arrays are hard in VBA).  |
| mdlExcel        | It has some method to look for in rows, columns, etc.                        |
| mdlFiles        | It has some methods to open dialog box for files, dirs, delete files, etc.   |
| mdlInternet     | It has methods to make requests to internet and URL encode strings.          |
| mdlStrings      | It has some methods to search case insensitive in strings and execute regex. |
| mdlUtils        | It has some methods like get an array number or wait.                        |
| mdlWord         | It lets you interact with some methods in Word Apps.                         |

## Example

```vba
Public Sub showNextEmptyRow()
  Dim nextEmptyRow as Single: nextEmptyRow = mdlExcel.getNextEmptyRow(WorkSheet1, 1, 1, 100)
  MsgBox nextEmptyRow
End Sub
```

# License

MIT @ [Matias Puig](https://www.github.com/matipuig)
