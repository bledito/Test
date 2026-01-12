from GetData import *

client_nm = 'Kenvue'
email_folder_path = 'Sofia>JnJ'
source_folder = r'\\10.7.16.12\ssis\Data Sources\CS JnJ'
source_folder_power_query = source_folder + '\\Power Query'
source_folder_excel = source_folder + '\\Excel'
archive_folder = source_folder + '\\Archive'

################################################ Call ##################################################################
Email(client_nm, 'Call',
      'JnJ ConsumerCare Berlin', 'xlsx', source_folder_excel + '\\Call', email_folder_path=email_folder_path,
      arch_folder_path=archive_folder + '\\Call', excel_to_csv=False)
SharedFolder(client_nm, 'Call',
             source_folder_power_query, 'Call.xlsx', source_folder, 't_calls.csv',
             excel_sheet='Call', power_query_name='Call',
             power_query_datetime_columns_formats='Date > dd/mm/yyyy; Time Period (15 min) > hh:mm:ss',
             power_query_source_path_to_delete=source_folder_excel + '\\Call')

############################################ Calls Details #############################################################
# ZIP extracted from email, CSVs unzipped, renamed and archived
Email(client_nm, 'Calls Details',
      'JnJ Call Details Daily', 'zip', source_folder, email_folder_path=email_folder_path,
      arch_folder_path=archive_folder + '\\Call Details')

################################################ CSAT ##################################################################
# Excel file extracted from email, converted to CSV and archived
Email(client_nm, 'CSAT',
      '[Kenvue] EMEA Survey Data', 'xlsx', source_folder, email_folder_path=email_folder_path,
      save_as_file_name='t_csat',
      arch_folder_path=archive_folder + '\\CSAT', add_time_to_archive=True,
      add_orig_excel_file_name_col=True, excel_sheet='Sheet 1', data_types='Created Date:datetime64[ns]')

############################################## CSAT Voice ##############################################################
# Excel file extracted from email, converted to CSV and archived
Email(client_nm, 'CSAT Voice',
      'Kenvue Task Details Task', 'xlsx', source_folder, email_folder_path=email_folder_path,
      file_name_contains_str='csat_per_row', save_as_file_name='t_csat_voice',
      arch_folder_path=archive_folder + '\\CSAT Voice',
      add_orig_excel_file_name_col=True, excel_sheet='Sheet1', data_types='ani:str;ucid:str')
