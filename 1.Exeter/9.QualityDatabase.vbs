Option Explicit

dim objExcel, objWorkbook
dim accessApp
dim MyApp, MyItem
dim wshNetwork, strUserName
dim dtmEnd, dtmStart, dtTime
dim check 

dtmStart=Now
Set wshNetwork = CreateObject( "WScript.Network" )
strUserName = wshNetwork.UserName

on error Resume next

set accessApp = createObject("Access.Application")
accessApp.visible = false

accessApp.UserControl = false

accessApp.OpenCurrentDataBase("\\gb998k12fps01\Client$\Flybe\QM Loggers\DO NOT TOUCH\QM Logger Database.accdb")
accessApp.Run "MergeDatabases"
accessApp.quit
set accessApp=nothing

set accessApp = createObject("Access.Application")
accessApp.visible = false

accessApp.UserControl = false
accessApp.OpenCurrentDataBase("\\rs127k16fps01\Flybe$\Data_Flybe\Private\QM logger\Do Not Touch\QM Logger Database.accdb")
accessApp.Run "MergeDatabases"
accessApp.quit
set accessApp=nothing


Set MyApp = CreateObject("Outlook.Application")
Set MyItem = MyApp.CreateItem(0)
With MyItem
    .To = "crt_uk@sitel.com"
    .ReadReceiptRequested = False
	if err.number>1 then
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
accessApp.Run "DailyMacro", Check, Wscript.ScriptName, strUserName, dtTime
accessApp.quit

'WScript.Echo "Finished."
 WScript.Quit