'if WScript.Arguments.Count < 1 Then
   ' WScript.Echo "Please specify the source  files. Usage: ExcelToCsv <xls/xlsx source file>"
  '  Wscript.Quit
'End If

csv_format = 62

dim oldCSV
oldCSV = "\\10.7.16.12\ssis\Data Sources\WFM\CAP\Latest\SYNTHESYS - Consolidated GA.csv"
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFile(oldCSV)

src_file = "\\10.7.16.12\ssis\Data Sources\WFM\CAP\Latest\SYNTHESYS - Consolidated GA.xlsb"


dest_file = Replace(Replace(src_file,".xlsb",".csv"),".xls",".csv")

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)

oBook.SaveAs dest_file, csv_format

oBook.Close False
oExcel.Quit