<div align="center">

## Read Excel Spread Sheet


</div>

### Description

The purpose of the following code is to provide you with a series of prototype functions to open and retreive data from a MS Excel spread sheet. The following code should be inserted into a new module named, for example, "modReadExcel". Passing variables will set the Excel File Name to open, the active Excel Sheet, recover data (data is returned as a string variable), close and exit Excel and clear the memory. These Prototype function simplify the entire process and gives your program(s) less coding or what I refer to as Clutter.

<br><br>

This code provides you with the basics of opening and reading an excel spreadsheet. I will be updating it in the future with the more advanced features if and when I encounter them.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan MacFarlane](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-macfarlane.md)
**Level**          |Intermediate
**User Rating**    |4.4 (75 globes from 17 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-macfarlane-read-excel-spread-sheet__1-14325/archive/master.zip)





### Source Code

<font color="grey">
'-------------------------------------------------<br>'
<br>'Excel Spread Sheet Read Prototype Functions
<br>'
<br>'---------------------------------------------<br>'
<br>' By Duncan MacFarlane
<br>' MacFarlane System Solutions
<br>' A Privately owned business operated <br>'  from personal residence
<br>'
<br>' Copyright MacFarlane System Solutions <br>'  2001
<br>'
<br>'---------------------------------------------<br>'
<br>' The following functions simplify <br>'  the process of opening,
<br>'  retrieving, closing, exiting
<br>'  Excel and clearing the memory of <br>'  the excel objects.
<br>'
<br>'---------------------------------------------<br>'
<br>' The Syntax of the following functions <br>'  are as follows:
<br>'
<br>'  excelFile([String - File Name Including Full Path])
<br>'  Sets the current file to open
<br>' excelPassword([String - Excel <br>'  Read Only Password], [String - <br>'  Excel Write Password]
<br>'  if no password is used on the <br>'  file discard the use of this <br>'  function
<br>' openExcelFile
<br>'  No variables are passed, opens <br>'  file set by excelFile function
<br>' setActiveSheet([Integer - Sheet <br>'  number of sheet to read from, <br>'  starting from 1]
<br>'  Sets the active sheet to read <br>'  from
<br>'  [String - Data input returned] = <br>' readExcel([Integer - Row], <br>'  [Integer - Column])
<br>'  Reads the content of a cell and <br>'  returns the data to the calling <br>'  location
<br>' closeExcelFile
<br>'  Closes the active Excel File
<br>' exitExcel
<br>'  Exits MS Excel
<br>' clearExcelObjects
<br>'  Clear the memory of the Excel <br>'  Application objects
<br>'---------------------------------------------</font>
<br><br>
<font color="blue">Dim</font> <font color="red">excelFileName</font> <font color="blue">As String</font>
<br>
<font color="blue">Dim</font> <font color="red">readPassword</font> <font color="blue">As String</font>
<br>
<font color="blue">Dim</font> <font color="red"> writePassword</font> <font color="blue">As String</font>
<br>
<font color="blue">Dim</font> <font color="red">msExcelApp</font> <font color="blue">As</font> <font color="red">Excel.Application</font>
<br>
<font color="blue">Dim</font> <font color="red">msExcelWorkbook</font> <font color="blue">As</font> <font color="red">Excel.Workbook</font>
<br>
<font color="blue">Dim</font> <font color="red">msExcelWorksheet</font> <font color="blue">As</font> <font color="red">Excel.Worksheet</font>
<br><br>
<font color="blue">Public Function </font> <font color="red">excelFile(fileName <font color="blue">As String</font><font color="red">)</font>
<br>
  <font color="blue">Let</font> <font color="red">excelFileName = fileName</font>
<br>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function</font> <font color="red">excelPassword(rdExcel</font> <font color="blue">As String</font><font color="red">, wtExcel</font> <font color="blue">As String</font><font color="red">)</font>
  <font color="blue">Let</font> <font color="red">readPassword = rdExcel</font<
<br>
  <font color="blue">Let</font> <font color="red">writePassword = rdExcel</font>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function</font> <font color="red">openExcelFile()</font>
<br>
  <font color="blue">Set</font> <font color="red">msExcelApp = GetObject(</font><font color="blue">""</font><font color="red">,</font> <font color="blue">"excel.application"</font><font color="red">)</font>
<br>
  <font color="red">msExcelApp.Visible =</font> <font color="blue">False</font>
<br>
  <font color="blue">If</font> <font color="red">readPassword =</font> <font color="blue">"" And</font> <font color="red">writePassword =</font> <font color="blue">"" Then</font>
<br>
    <font color="blue">Set</font> <font color="red">msExcelWorkbook = Excel.Workbooks.Open(excelFileName)</font>
<br>
  <font color="blue">Else</font>
<br>
    <font color="blue">Set</font> <font color="red">msExcelWorkbook = Excel.Workbooks.Open(excelFileName, , , , readPassword, writePassword)</font>
<br>
  <font color="blue">End If</font>
<br>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function</font> <font color="red">setActiveSheet(excelSheet <font color="blue">As Integer</font><font color="red">)</font>
<br>
  <font color="blue">Set</font> <font color="red">msExcelWorksheet = msExcelWorkbook.Worksheets.Item(excelSheet)</font>
<br>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function</font> <font color="red">readExcel(Row</font> <font color="blue">As Integer</font><font color="red">, Col</font> <font color="blue">As Integer</font><font color="red">)</font> <font color="blue">As String</font>
<br>
  <font color="red">readExcel = msExcelWorksheet.Cells(Row, Col)</font>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function,</font> <font color="red">closeExcelFile()</font>
<br>
  <font color="red">msExcelWorkbook.Close</font>
<br>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function</font> <font color="red">exitExcel()</font>
<br>
  <font color="red">msExcelApp.Quit</font>
<font color="blue">End Function</font>
<br><br>
<font color="blue">Public Function</font> <font color="red">clearExcelObjects()</font>
  <font color="blue">Set</font> <font color="red">msExcelWorksheet =</font> <font color="blue">Nothing</font>
<br>
  <font color="blue">Set</font> <font color="red">msExcelWorkbook =</font> <font color="blue">Nothing</font>
<br>
  <font color="blue">Set</font> <font color="red">msExcelApp =</font> <font color="blue">Nothing</font>
<br>
<font color="blue">End Function</font>

