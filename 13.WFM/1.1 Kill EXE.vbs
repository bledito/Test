Option Explicit

Dim ProcessToKill, WMI, Process, n_of_processes

ProcessToKill = "EXCEL.EXE"
n_of_processes = 0

Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT * FROM Win32_Process")

For Each Process In WMI
    If LCase(Process.Name) = LCase(ProcessToKill) Then
        Process.Terminate
        n_of_processes = n_of_processes + 1
    End If
Next

'If n_of_processes = 0 Then
 '   MsgBox "No Excel processes running.", vbInformation, "Script Result"
'Else
    'MsgBox "Script has run successfully. Closed " & n_of_processes & " Excel processes.", vbInformation, "Script Result"
'End If
