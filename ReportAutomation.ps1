$SQLServer = "EU999K16SSCDB01"
$SQLDBName = "CRTDB"
#$ErrorActionPreference variable: SilentlyContinue, Continue, Stop, and Inquire.
$ErrorActionPreference = "SilentlyContinue"

## Check who is running the report from s_a_responsible table
$SqlQuery = "Select [Responsible] FROM  [CRTDB].[dbo].[s_a_responsible]"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
 
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
 
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
 
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$usercheck=$dataset.Tables[0]
$environ=$env:username
$SqlConnection.Close()

$counter=0
#Break if not the responsible user
foreach ($Row in $dataset.Tables[0].Rows)
{ 
  if (($row[0].Contains($environ))) {$counter=$counter+1}
}
if ($counter -ne 1) {stop-process -Id $PID}

## Check for available reports
$SqlQuery = "Select [Report Name],Site, [Report iD] FROM  [CRTDB].[dbo].[v_z_update_report_status] where Status='Ready'"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
 
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
 
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
 
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
 
$SqlConnection.Close()

#Format date for the where clause
$a=(get-date)
$a=$a.ToString("yyyy-MM-dd")
$a="'$a'"

#Looping through each available report
foreach ($Row in $dataset.Tables[0].Rows)
{ 
  $ReadyReport = $($Row[0]) + " " + $($Row[1])
  $ReportID=$row[2]
  $ReportCheck=0
   
#Taking Filepath, Macro from the ReportList powershell script *** ADD NEW REPORTS THERE ***

  $ReturnArray = Powershell -noprofile -executionpolicy bypass -file "D:\Reports\1.PS\ReportList.ps1" $ReadyReport
  $File = $ReturnArray[0]
  $Macro = $ReturnArray[1]
  $ReportCheck = $ReturnArray[2] 


#Check if the report is available
      if ($reportcheck -eq 1) 
      {
             $StartTime = $(get-date)
             $excel = new-object -comobject excel.application
             $workbook = $excel.workbooks.open($File)
             $excel.Run($Macro,$ReportID)

             $excelprocess = ((get-process excel | select MainWindowTitle, ID, StartTime | Sort StartTime)[-1]).Id
             Stop-Process -Id $excelprocess

#Check if there was an error with the report
             $SqlQuery="SELECT [Error Count] FROM [dbo].[s_report_update_status] WHERE [DATE]=$a and [Report ID]=$ReportID"
             $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
             $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
  
             $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
             $SqlCmd.CommandText = $SqlQuery
             $SqlCmd.Connection = $SqlConnection
 
             $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
             $SqlAdapter.SelectCommand = $SqlCmd 
 
             $DataSet = New-Object System.Data.DataSet
             $SqlAdapter.Fill($DataSet)   | Out-Null  
     
             $SqlConnection.Close()

             $elapsedTime = $(get-date) - $StartTime
             $elapsedtime="'$elapsedtime'"

             $CompletionTime=(get-date)
             $CompletionTime=$CompletionTime.ToString("hh:mm:ss")
             $CompletionTime="'$CompletionTime'"


#Check if there were any error messages in the macro and rerun it 3 times
             if (!$Dataset.Tables[0].Rows[0].ItemArray.Contains(0))
             {
                DO
                {  
                    $excel = new-object -comobject excel.application
                    $workbook = $excel.workbooks.open($File)
                    $excel.Run($Macro,$ReportID)

                    $excelprocess = ((get-process excel | select MainWindowTitle, ID, StartTime | Sort StartTime)[-1]).Id
                    Stop-Process -Id $excelprocess
                    $SqlQuery="SELECT [Error Count] FROM [dbo].[s_report_update_status] WHERE [DATE]=$a and [Report ID]=$ReportID"
                    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
                    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True" 
                    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
                    $SqlCmd.CommandText = $SqlQuery
                    $SqlCmd.Connection = $SqlConnection 
                    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
                    $SqlAdapter.SelectCommand = $SqlCmd
                    $DataSet = New-Object System.Data.DataSet
                    $SqlAdapter.Fill($DataSet)    | Out-Null  
                    $SqlConnection.Close()
                } WHILE ($Dataset.Tables[0].Rows[0].ItemArray -gt 0 -and $Dataset.Tables[0].Rows[0].ItemArray -lt 4)

#Check for error status after 3 runs          
                if ($Dataset.Tables[0].Rows[0].ItemArray -gt 3)
                {
                    $SqlQuery = "update [s_report_update_status] set [Status]='Error',[Update Time]=$elapsedTime, [Completion Time]=$CompletionTime  where [report ID]=$ReportID and [date]=$a"
                    write-host " "
                    write-host "The " $ReadyReport " has found an error" -ForegroundColor "Red"
                    write-host " "
                }
                else 
                {
                    $SqlQuery = "update [s_report_update_status] set [Status]='Completed',[Update Time]=$elapsedTime, [Completion Time]=$CompletionTime  where [report ID]=$ReportID and [date]=$a"
                    write-host " "
                    write-host "The " $ReadyReport " has been successful" -ForegroundColor "Green"
                    write-host " "
                }
            }
            else
            {
            $SqlQuery = "update [s_report_update_status] set [Status]='Completed',[Update Time]=$elapsedTime, [Completion Time]=$CompletionTime  where [report ID]=$ReportID and [date]=$a"
            write-host " "
            write-host "The " $ReadyReport " has been successful" -ForegroundColor "Green"
            write-host " "
            }

            $con = new-object "System.data.sqlclient.SQLconnection"
            $con.ConnectionString =(“Server = $SQLServer; Database = $SQLDBName; Integrated Security = True”)
            $con.open()

            $sqlcmd = new-object "System.data.sqlclient.sqlcommand"
            $sqlcmd.connection = $con

            $sqlcmd.CommandText = $SqlQuery 
            $sqlcmd.ExecuteNonQuery()  | Out-Null         
      }

#Fail if macro and filepath have not been updated in the ReportList script
      else 
      {
            write-host " "
            write-host "Unavailable report data for" $ReadyReport -ForegroundColor "Red"
            write-host " "

            $SqlQuery = "update [s_report_update_status] set [Status]='Failed' where [report ID]=$ReportID and [date]=$a"

            $con = new-object "System.data.sqlclient.SQLconnection"
            $con.ConnectionString =(“Server = $SQLServer; Database = $SQLDBName; Integrated Security = True”)
            $con.open()

            $sqlcmd = new-object "System.data.sqlclient.sqlcommand"
            $sqlcmd.connection = $con

            $sqlcmd.CommandText = $SqlQuery
            $sqlcmd.ExecuteNonQuery() | Out-Null 
      }
      $con.close()
}

#stop-process -Id $PID