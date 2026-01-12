$SQLServer = "EU999K16SSCDB01"
$SQLDBName = "CRTDB"
$arraychoice=@()

#$ErrorActionPreference variable: SilentlyContinue, Continue, Stop, and Inquire.
$ErrorActionPreference = "SilentlyContinue"

clear-host

## Check for available reports
$SqlQuery = "Select [Report Name],Site, [Report iD] FROM  [CRTDB].[dbo].[s_reports] where Active='Yes'"

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

clear-host
write-host "Report IDs:" -ForegroundColor Yellow

#Write available reports in console
foreach ($Row in $dataset.Tables[0].Rows)
{ if($Row[2] -lt 10){$space=" "} else {$space=""} 
  $ReadyReport = $($Row[0]) + " " + $($Row[1])
  $ReportID=$row[2]
  write-host $ReportID $space $ReadyReport -ForegroundColor "Yellow"
}

#Call Input Box
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Choose a report to rerun'
$msg   = 'Enter Report ID (you can add multiple report IDs, separated by comma. Find the list in the console.By typing "t" you can rerun by table ID or you can run all reports by typing *'

$choice = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
if  (([string]::IsNullOrEmpty($choice))) {Break}

#If rerun by Table ID was chosen                     ******* Table PART ********
if ($choice -eq "t")
{
#Write Table IDs and Names in console
    $SqlQuery = "Select distinct a.[Table ID],[Table Name] 
                    FROM  
                    (
	                    (select * from [CRTDB].[dbo].[s_tables]) as a 
	                    left join 
	                    (
	                    select b.[table id], b.[Report ID] 
	                    from
	                    (
		                    (select * from [CRTDB].[dbo].[s_reports_mapping]) as b
		                    inner join 
		                    (select * from [dbo].[s_reports] where [Active]='Yes') as c
		                    on b.[Report ID]=c.[Report ID]
	                    )
	
	                    ) as e on a.[Table Id]=e.[Table ID]
                    )
                    where e.[Report ID] is not null"


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

    clear-host
    write-host "Table IDs:" -ForegroundColor "Yellow"

    foreach ($Row in $Dataset.Tables[0].Rows)
    { 
        if($row[0] -lt 10){$space=" "} else {$space=""}
        write-host $Row[0] $space $row[1] -ForegroundColor "Yellow"
    }

#Inputbox for Table ID
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    $title = 'Choose Table ID'
    $msg   = 'Enter Table ID (you can add multiple table IDs, separated by comma. Find the list in the console.'

    $choice = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    if  (([string]::IsNullOrEmpty($choice))) {Break}

#Select available Report IDs for rerun
    $SqlQuery = "Select distinct e.[Report ID] 
                FROM  
                (
	                (select * from [CRTDB].[dbo].[s_tables]) as a 
	                left join 
	                (
	                select b.[table id], b.[Report ID] 
	                from
	                (
		                (select * from [CRTDB].[dbo].[s_reports_mapping]) as b
		                inner join 
		                (select * from [dbo].[s_reports] where [Active]='Yes') as c
		                on b.[Report ID]=c.[Report ID]
	                )
	
	                ) as e on a.[Table Id]=e.[Table ID]
                )
                where e.[Report ID] is not null and a.[table id] in ($choice)"


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

#Put all reports from Table inputbox into array
    foreach ($Row in $Dataset.Tables[0].Rows)
    {
    $arraychoice+=$Row[0]
    }

#Load Report ID $Dataset
    $SqlQuery = "Select [Report Name],Site, [Report iD] FROM  [CRTDB].[dbo].[s_reports] where Active='Yes'"
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

}
#Else if Report ID was initially chosen                           ******* Table PART END********
else
{
#Put all reports from Report inputbox into array
    $index = $choice.IndexOf(",")
    if ($index -ne -1) 
    {
        do 
        {
        $arraychoice+=$choice.substring(0,$index)
        $textLen= $choice.length
        $arrayLen=$choice.substring(0,$index).Length+1
        $startpos=$index+1
        $choice=$choice.Substring($startpos,$textLen-$arrayLen)
        $index = $choice.IndexOf(",")
        } while ($index -ne -1)
    }

    $arraychoice+=$choice
}

write-host " "
write-host "Loading..."

#Loop through the chosen reports 
foreach ($Row in $dataset.Tables[0].Rows)
{ 

  if (-not @($Row[2]| where {$arraychoice -notcontains $_}).Count -or $choice -eq "*")
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
}

$arraychoice = $null 
Write-host "All Done!" -ForegroundColor "Yellow"