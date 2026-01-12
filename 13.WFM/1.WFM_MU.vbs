Option Explicit

Dim objExcel, objWorkbook
Dim MyApp, MyItem
Dim wshNetwork, strUserName
Dim dtmStart, check

dtmStart = Now
Set wshNetwork = CreateObject("WScript.Network")
strUserName = wshNetwork.UserName

' Turn OFF silent error suppression
On Error GoTo 0

' Create Excel Application
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False            ' Prevent save prompts
objExcel.Visible = False                  ' Keep hidden

' --- OPEN WORKBOOK ---
Set objWorkbook = objExcel.Workbooks.Open("\\EU999K16SSCDB01\ssis\Data Sources\WFM\Mapping\MU Mapping\MU_mapping PQ Converter.xlsb")

' --- RUN MACROS ---
Dim macroError : macroError = False

On Error Resume Next     ' Only for controlled macro failure detection

objExcel.Run "refresh_connections"
If Err.Number <> 0 Then
    macroError = True
    Err.Clear
End If

WScript.Sleep 30000

objExcel.Run "refresh_connections"
If Err.Number <> 0 Then
    macroError = True
    Err.Clear
End If

WScript.Sleep 30000

objExcel.Run "save_csv"
If Err.Number <> 0 Then
    macroError = True
    Err.Clear
End If

On Error GoTo 0     ' Turn error handling back on

' --- CLOSE WORKBOOK ---
' Close without saving workbook changes (but macro save_csv handles external saving)
If Not objWorkbook Is Nothing Then
    objWorkbook.Close False
End If

' --- QUIT EXCEL ---
objExcel.Quit

' --- RELEASE COM OBJECTS IN CORRECT ORDER ---
Set objWorkbook = Nothing
Set objExcel = Nothing

' --- SEND EMAIL REPORT ---
Set MyApp = CreateObject("Outlook.Application")
Set MyItem = MyApp.CreateItem(0)

With MyItem
    .To = "crt_uk@sitel.com"
    .ReadReceiptRequested = False

    If macroError = True Then
        .HTMLBody = "Failed: One or more Excel macros failed. Please re-run the script."
        .Subject = "FAILED Script: WFM MU Mapping - " & WScript.ScriptName
    Else
        .HTMLBody = "Script has run successfully."
        .Subject = "Script: WFM MU Mapping - " & WScript.ScriptName
    End If
End With

MyItem.Send

WScript.Quit
