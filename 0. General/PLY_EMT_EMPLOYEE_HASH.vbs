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

accessApp.OpenCurrentDataBase("\\GB645K12FPS01\Administrative$\Operations\Shared\9. Databases DO NOT TOUCH\PLY Employee Details1.accdb")
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