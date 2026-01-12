Option Explicit
dim objExcel, objWorkbook
dim accessApp
dim MyApp, MyItem
dim wshNetwork, strUserName
dim dtmEnd, dtmStart, dtTime
dim check 

dtmStart = Now
Set wshNetwork = CreateObject( "WScript.Network" )
strUserName = wshNetwork.UserName

on error resume next

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\3. IOT\Breaks & IOT Report.xlsb")

objExcel.Application.Visible = false
objExcel.run "SendEmail"

objWorkbook.close False
objExcel.quit
set objWorkbook = Nothing
set objExcel = Nothing


Set MyApp = CreateObject("Outlook.Application")
Set MyItem = MyApp.CreateItem(0)
With MyItem
    .To = "crt_uk@sitel.com"
    .ReadReceiptRequested = False
   if err.number<>0 then
	.HTMLBody = "Failed: Re-run script"
    .Subject = "Failed Script: Exeter Status Report - " &  Wscript.ScriptName
	Check="Fail"
	else
    .HTMLBody = "Script has run successfully"
	.Subject = "Script: Exeter Status Report - " &  Wscript.ScriptName
	Check="OK"
	end if
End With
MyItem.Send

set accessApp = createObject("Access.Application")
accessApp.visible = false

accessApp.UserControl = false


dtmEnd = Now
dtTime=DateDiff("s", dtmStart, dtmEnd)

'WScript.Echo Wscript.ScriptName, strUserName, dtTime
accessApp.OpenCurrentDataBase("\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\1. Global\1.Databases\Tasks Database.accdb")
accessApp.Run "DailyMacro",Check, Wscript.ScriptName, strUserName, dtTime
accessApp.quit

'WScript.Echo "Finished."
 WScript.Quit