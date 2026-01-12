Option Explicit

dim objExcel, objWorkbook
dim accessApp
dim MyApp, MyItem
dim wshNetwork, strUserName
dim dtmEnd, dtmStart, dtTime

dtmStart=Now

set accessApp = createObject("Access.Application")
accessApp.visible = true

accessApp.UserControl = false

accessApp.OpenCurrentDataBase("\\gb998k12fps01\administrative$\Operations\Shared\Manager Skill Reports Source DO NOT TOUCH\EXE Employee Details.accdb")
accessApp.Run "Hash"
accessApp.quit
set accessApp=nothing

Set MyApp = CreateObject("Outlook.Application")
Set MyItem = MyApp.CreateItem(0)
With MyItem
    .To = "crt_uk@sitel.com"
    .Subject = "Script: Plymouth Status Report - " &  Wscript.ScriptName
    .ReadReceiptRequested = False
    .HTMLBody = "Script has run successfully"
End With
MyItem.Send

'WScript.Echo "Finished."
 WScript.Quit