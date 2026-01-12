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

accessApp.OpenCurrentDataBase("\\gb998k12fps01\administrative$\Operations\Restricted\John Lewis\5. OSA\Combined Login Tool\Database\HEMT Database DO NOT TOUCH.accdb")
accessApp.Run "update_sql_tables"
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

accessApp.quit

'WScript.Echo "Finished."
 WScript.Quit