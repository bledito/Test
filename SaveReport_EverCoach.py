from GetData import *

s3_folder_path = r'\\emeaits.emea.sitel-world.net\Prod\reports'
task_nm = 'EverCoach Upload'

# Mondays
if datetime.today().weekday() == 0:
    # Epson
    GenerateEverCoachReport('Epson', task_nm, database_nm='CS_Epson',
                            sql_query='SELECT * FROM [REP].[i_coach_agent]',
                            output_folder_path=s3_folder_path + '\\576TOT1',
                            output_file_name='Multiple CAPs Epson_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True)

    # ZF
    GenerateEverCoachReport('ZF', task_nm, database_nm='CS_ZF',
                            sql_query='SELECT * FROM [dbo].[v_i_coach_agent]',
                            output_folder_path=s3_folder_path + '\\11340T1',
                            output_file_name='CAP-15657 ZF_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True)

    # Oventrop
    GenerateEverCoachReport('Oventrop', task_nm, database_nm='CS_Oventrop',
                            sql_query='SELECT * FROM [dbo].[v_iCoach] WHERE [DENOMINATOR] <> 0',
                            output_folder_path=s3_folder_path + '\\11970T1',
                            output_file_name='CAP-17773 CAP-19303 Oventrop GmbH_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True)

    sql_query_file_date = "SELECT DATEADD(DAY, -DATEPART(WEEKDAY, GETDATE()) - 12, CAST(GETDATE() AS date)) AS 'STARTDATE'"
    # Kenvue Sofia
    GenerateEverCoachReport('Kenvue', task_nm + ' (Sofia)', database_nm='CS_QM_JnJ',
                            sql_query="SELECT * FROM [REP].[v_icoach_agent] WHERE [LOB] = 'CAP-8261:Kenvue - Social Media Support'",
                            output_folder_path=s3_folder_path + '\\481TOT1',
                            output_file_name='CAP-8261 Kenvue_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True,
                            sql_query_file_date=sql_query_file_date)

    # Kenvue Berlin
    GenerateEverCoachReport('Kenvue', task_nm + ' (Berlin)', database_nm='CS_QM_JnJ',
                            sql_query="SELECT * FROM [REP].[v_icoach_agent] WHERE [LOB] = 'CAP-7193:Kenvue -   Consumer Care Team'",
                            output_folder_path=s3_folder_path + '\\481TOT1',
                            output_file_name='CAP-7193 Kenvue_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True,
                            sql_query_file_date=sql_query_file_date)

    # Kenvue Athens
    GenerateEverCoachReport('Kenvue', task_nm + ' (Athens)', database_nm='CS_QM_JnJ',
                            sql_query="SELECT * FROM [REP].[v_icoach_agent] WHERE [LOB] = 'CAP-15241:Johnson & Johnson'",
                            output_folder_path=s3_folder_path + '\\481TOT1',
                            output_file_name='CAP-15241 Kenvue_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True,
                            sql_query_file_date=sql_query_file_date)
# Tuesdays
if datetime.today().weekday() == 1:
        # Medtronic
        GenerateEverCoachReport('Medtronic', task_nm, database_nm='CS_Medtronic',
                                sql_query='SELECT * FROM [dbo].[v_iCoach_Agent]',
                                output_folder_path=s3_folder_path + '\\061TOT1\\EverCoach',
                                output_file_name='Multiple CAPs Medtronic_%m%d%Y.xlsx',
                                datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True)

# Mondays and Wednesdays
if datetime.today().weekday() == 0 or datetime.today().weekday() == 2:
    # Klarmobil
    GenerateEverCoachReport('Klarmobil', task_nm, database_nm='CS_KM',
                            sql_query='SELECT * FROM [REP].[v_iCoach_Upload] WHERE [DENOMINATOR] <> 0 AND [NUMERATOR] IS NOT NULL',
                            output_folder_path=s3_folder_path + '\\049TOT1',
                            output_file_name='CAP-7078 Klarmobil_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True)

    # Freenet
    GenerateEverCoachReport('Freenet', task_nm, database_nm='CS_MD',
                            sql_query='SELECT * FROM [REP].[v_2020_upload] WHERE [DENOMINATOR] <> 0 AND [NUMERATOR] IS NOT NULL',
                            output_folder_path=s3_folder_path + '\\105TOT1',
                            output_file_name='CAP-5965 Mobilcom Debitel_%m%d%Y.xlsx',
                            datetime_format='dd/mm/yyyy', start_row=1, add_date_to_file_name=True)

# Thursdays
if datetime.today().weekday() == 3:
    # RingCentral
    GenerateEverCoachReport('RingCentral', task_nm, database_nm='CS_RC',
                            sql_query='SELECT * FROM [REP].[v_icoach]',
                            output_folder_path=s3_folder_path + '\\10220T1',
                            output_file_name='CAP-13726 RingCentral_%m%d%Y.xlsx',
                            start_row=1, add_date_to_file_name=True)

