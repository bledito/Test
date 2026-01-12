
ProcessToKill = "EXCEL.EXE"
Set WMI=GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery ("select * from Win32_Process")
For Each Process in WMI 
               If process.name = ProcessToKill Then
                              Process.terminate
               End If
Next
