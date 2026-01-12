param
(
$ReadyReport
)
   
$ReturnArray=@()

#Add new reports below: Filepath,Macro Name, Report Check (always 1)

     if ($ReadyReport -like "*Absence Dashboard Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\1. Absence\Absence Report Exeter SQL.xlsb"
             $ReturnArray +="Main_macro_SQL"
             $ReturnArray += 1
      }
  
      elseif ($ReadyReport -like "*Absence Dashboard Plymouth*") 
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\6. Plymouth\3. Standard reports\0. Generic\1. Absence\Absence Report Plymouth SQL.xlsb"
             $ReturnArray +="Main_macro_SQL"
             $ReturnArray +=1
      }

      elseif ($ReadyReport -like "*Telephony Dashboard Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\2. Telephony\Master Telephony sql.xlsb"
             $ReturnArray +="Main_macro_SQL"
             $ReturnArray += 1
      }    

      elseif ($ReadyReport -like "*Heathrow MI Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\1. Heathrow\Heathrow MI Master SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }    
      
      elseif ($ReadyReport -like "*Voyage Prive MI Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\4. Voyage Prive\Voyage Prive MI - SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*John Lewis IVR MI Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\5. John Lewis Visual IVR\John Lewis Visual IVR.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*John Lewis EHT MI Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\1. John Lewis EHT After Sales\John Lewis EHT After Sales SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*John Lewis HomeSolutions Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\6. John Lewis Home Solutions\John Lewis Home Solutions MI SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*John Lewis TechnicalSupport Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\4. John Lewis Technical Support\John Lewis Techincal Support MI Master.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*Toshiba MI Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\2. Toshiba\Toshiba MI Master SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*Gradbay Telephony Performance Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\6. Gradbay Telephony\Gradbay Telephony Performance.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*John Lewis CustomerService Plymouth*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\6. Plymouth\3. Standard reports\1. John Lewis\2. John Lewis Customer Service\John Lewis Customer Service SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*SIOX Master SQL Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\7. SIOX\SIOX Master SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 

      elseif ($ReadyReport -like "*JnJ MI Exeter*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\11. J&J\J&J MI Draft SQL.xlsm"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*Login-logout*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\4. Login-Logout\EXE PLY Login-Logout Report SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*SIOX Comments Master SQL*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\0. Generic\7. SIOX\SIOX Comments Master SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*Hunter MI Report*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\4. Coventry\3. Standard reports\5 HunterBoot\Hunter MI Report .xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*Hunter Agent Snapshot*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\4. Coventry\3. Standard reports\5 HunterBoot\Agent Snapshot .xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
      elseif ($ReadyReport -like "*Flybe MI*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\5. Flybe\1. Flybe MI Report\Flybe MI Master SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*Track and Trace MI Ply*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\13. NHS\NHS - Track and Trace MI Master SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*Track and Trace MI - External*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\13. NHS\NHS - Track and Trace MI Master SQL - External.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*NHS - Track and Trace Synergy SQL*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\13. NHS\NHS - Track and Trace Synergy SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*Tier 2 & 3 DHSC*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\13. NHS\Tier 2 & 3 DHSC SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*NHS - Track and Trace Synergy - Tier 2*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\13. NHS\NHS - Track and Trace Synergy - Tier 2 SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
       elseif ($ReadyReport -like "*JL Back Office Productivity*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\9.JL Back Office Productivity\JL Back Office Productivity SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*Absence Report Stratford*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\20.Stratford\4.Standard Reports\0.Generic\1.Absence Report\Absence Report Stratford SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*JL OB Number Dialed*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\12. JL OB Number Dialed\JL OB Number Dialed SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*John Lewis TCD*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\2. Exeter\4. Standard reports\3. John Lewis\8. TCD Report\John Lewis TCD SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      }
      elseif ($ReadyReport -like "*Waitrose MI*")
      {
             $ReturnArray += "\\BG279K12FPS01\Administrative$\Reporting\Data_Reporting\Private\1. Projects\6. Plymouth\3. Standard reports\3. Waitrose\1. Waitrose MI\Waitrose MI SQL.xlsb"
             $ReturnArray +="Main_macro_sql"
             $ReturnArray += 1
      } 
return $ReturnArray