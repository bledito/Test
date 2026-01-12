from GetData import *
import dateutil.relativedelta

server_nm = 'EU999K16SSCDB01'
database_nm = 'CS_Lexmark'
output_path = r'\\10.17.25.15\Administrative$\Reporting\Data_Reporting\Private\1. Projects\8. Sofia\4. Standard Reports\Lexmark'


def GenerateReport(sql_query, subject_file_nm, recipients, body):
    output_filepath = output_path + '\\' + subject_file_nm + '.xlsx'
    df = DataFrameFromSQL(server_nm, database_nm, sql_query)
    DataFrameToExcel(df, output_filepath)
    SendEmail(subject_file_nm, body, recipients, attachment_filepath=output_filepath)


# checks whether to send the MTD report
if datetime.today().weekday() == 0 and datetime.today().day != 1:  # weekdays are from 0 to 6; on Monday, 1st there won't be data
    current_month_str = datetime.today().strftime('%B %Y MTD')
    GenerateReport(sql_query='SELECT * FROM [CS_Lexmark].[dbo].[v_IVR_CSAT_and_Telephony_MTD]',
                   subject_file_nm=('Lexmark IVR CSAT Report ' + current_month_str),
                   recipients='maryjane.suratos@foundever.com; rieza.espana@foundever.com; Gemma.Marino@foundever.com; Lenie.Gervero@foundever.com',
                   body=('Hello,<br></br><br>Attached you will find Lexmark IVR CSAT Report for current month.</br><br></br><br>' +
                         'Kind Regards,</br><br>The Reporting Team</br>')
                   )

# checks whether to send the Previous Month report
if datetime.today().day == 1:
    previous_month_str = (datetime.today() - dateutil.relativedelta.relativedelta(months=1)).strftime('%B %Y')
    GenerateReport(sql_query='SELECT * FROM [CS_Lexmark].[dbo].[v_IVR_CSAT_and_Telephony_Previous_Month]',
                   subject_file_nm=('Lexmark IVR CSAT Monthly Report - ' + previous_month_str),
                   recipients='maryjane.suratos@foundever.com; rieza.espana@foundever.com; Gemma.Marino@foundever.com; Lenie.Gervero@foundever.com',
                   body=('Hello,<br></br><br>Attached you will find Lexmark IVR CSAT Report for previous month.</br><br></br><br>' +
                         'Kind Regards,</br><br>The Reporting Team</br>')
                   )
