import os
import traceback
import win32com.client
import zipfile
import time
from datetime import datetime
import dateutil.relativedelta
import pandas as pd
import pyxlsb
import shutil
import pyodbc

# Python encodings: https://docs.python.org/3/library/codecs.html#standard-encodings
# Python date formatting strings: https://www.geeksforgeeks.org/python/how-to-format-date-using-strftime-in-python/


def RunExcelVBA(excel_filepath, vba_macro_name, save_after_close=True):
    excel_app = win32com.client.DispatchEx("Excel.Application")
    excel_wb = excel_app.Workbooks.Open(excel_filepath)
    excel_app.Visible = False
    excel_app.Run(vba_macro_name)
    excel_wb.Close(save_after_close)
    excel_app.Quit()


def DataFrameFromSQL(server_nm, database_nm, sql_query, date_columns=None, date_only_columns=None, data_types=None):
    # DESCRIPTION:
    # Returns pandas dataframe from query from SQL Server
    #
    # PARAMETERS:
    # server_nm: name of the server which will be queried
    # database_nm: name of the database which will be queried
    # sql query: SQL query which will be read
    # date_columns (optional): list of column names to parse as dates
    # date_only_columns (optional): list of column names which will be converted to dates (without time)
    #                               after the query is read
    # data_types (optional): list of column names and data types with which explicitly to set the data types
    #                        in the pandas dataframe

    conn = pyodbc.connect('DRIVER={SQL Server}' +
                          ';SERVER=' + server_nm +
                          ';DATABASE=' + database_nm +
                          ';Trusted_Connection=yes')
    df = pd.read_sql(sql_query, conn, parse_dates=date_columns, dtype=data_types)
    if date_only_columns is not None:
        for date_only_column in date_only_columns:
            df[date_only_column] = df[date_only_column].dt.date
    conn.close()
    return df


def DataFrameToExcel(dataframe, output_filepath, output_sheet_nm='Sheet1', date_format=None, datetime_format=None,
                     decimal_places='', column_width='', freeze_headers=False, auto_filter=False, start_row=1,
                     additional_dataframe=None, additional_dataframe_start_row=1):
    # DESCRIPTION:
    # Writes pandas dataframe to Excel file;
    # optionally writes additional dataframe in the same sheet but on different starting row (made for EverCoach)
    #
    # PARAMETERS:
    # dataframe: pandas dataframe object which will be written to the Excel file
    # output_filepath: file path of the Excel file
    # output_sheet_nm (optional): the name of the sheet in Excel where the data will be written;
    #                             if not set, the sheet will be 'Sheet1'
    # date_format (optional): the format string for the date objects written in the Excel file
    # datetime_format (optional): the format string for the datetime objects written in the Excel file
    # decimal_places (optional): the decimal places for the floating point numbers in the Excel file
    # column_width (optional): the size of the width of the columns in the Excel file;
    #                          if not set, the default size will be applied
    # freeze_headers (optional): indicates whether to freeze the Excel sheet where the header row is;
    #                            depends on the value of the 'start_row' where the header row will be;
    #                            the default is False (no freezing will be done)
    # auto_filter (optional): indicates whether to set auto filter on the header row;
    #                         depends on the value of the 'start_row' where the header row will be;
    #                         the default is False (no auto filter will be applied)
    # start_row (optional): the start row of the data (where the headers will be) which will be written
    #                       in the Excel sheet; if not set, the headers will be on the first row
    # additional_dataframe (optional): pandas dataframe object to be written on the same sheet as the main dataframe
    #                                  (designed for EverCoach reports);
    #                                  if not set, no additional data will be written in the Excel file
    # additional_dataframe_start_row (optional): the start row of the 'additional_dataframe' which will be written
    #                                            in the Excel sheet; it will be written after the main dataframe so
    #                                            make sure both don't overlap;
    #                                            if not set the start row will be the first row

    float_format = None
    if decimal_places != '':
        float_format = '%.' + str(decimal_places) + 'f'

    writer = pd.ExcelWriter(output_filepath, date_format=date_format, datetime_format=datetime_format)
    dataframe.to_excel(writer, sheet_name=output_sheet_nm, index=False, float_format=float_format, startrow=start_row-1)
    if additional_dataframe is not None:
        additional_dataframe.to_excel(writer, sheet_name=output_sheet_nm, index=False, float_format=float_format,
                                      startrow=additional_dataframe_start_row-1, header=False)

    output_sheet = writer.sheets[output_sheet_nm]
    if column_width != '':
        output_sheet.set_column(0, len(dataframe.columns) - 1, column_width)
    if freeze_headers:
        output_sheet.freeze_panes(start_row, 0)
    if auto_filter:
        output_sheet.autofilter(start_row-1, 0, start_row-1, len(dataframe.columns) - 1)
    writer.close()


def SendEmail(subject, body, recipients_to, recipients_cc='', recipients_bcc='', attachment_filepath=''):
    outlook = win32com.client.Dispatch('Outlook.Application')
    outlook_email = outlook.CreateItem(0)
    outlook_email.Subject = subject
    outlook_email.HTMLBody = body
    outlook_email.To = recipients_to
    outlook_email.CC = recipients_cc
    outlook_email.BCC = recipients_bcc
    if attachment_filepath != '':
        outlook_email.Attachments.Add(attachment_filepath)
    outlook_email.Send()


def GenerateEverCoachReport(project_nm, report_nm,
                            database_nm, sql_query, output_folder_path, output_file_name,
                            evercoach_parameters_str=None,
                            date_format=None, datetime_format=None, start_row=3, evercoach_parameters_start_row=1,
                            freeze_headers=False, auto_filter=False,
                            add_date_to_file_name=False, sql_query_file_date=None):
    current_action = 'read from SQL server'
    try:
        server_nm = 'EU999K16SSCDB01'
        df = DataFrameFromSQL(server_nm, database_nm, sql_query)
        if add_date_to_file_name:
            if sql_query_file_date is None:
                sql_query_file_date = 'SELECT CAST(MIN([STARTDATE]) AS date) STARTDATE FROM (' + sql_query + ') T'
            df_start_dt = DataFrameFromSQL(server_nm, database_nm, sql_query_file_date, date_columns=['STARTDATE'])
            report_start_dt = df_start_dt.iloc[0, 0]
            output_file_name = report_start_dt.strftime(output_file_name)
        output_filepath = os.path.join(output_folder_path, output_file_name)

        current_action = 'save as Excel'
        df_evercoach_parameters=None
        if evercoach_parameters_str is not None:
            evercoach_parameters = [row.split('=') for row in evercoach_parameters_str.split(';')]
            df_evercoach_parameters = pd.DataFrame(evercoach_parameters)
        DataFrameToExcel(df, output_filepath, date_format=date_format, datetime_format=datetime_format,
                         start_row=start_row, additional_dataframe=df_evercoach_parameters,
                         additional_dataframe_start_row=evercoach_parameters_start_row,
                         freeze_headers=freeze_headers, auto_filter=auto_filter)

        current_action = 'EverCoach report saved in destination folder'
        Log([str(project_nm), str(report_nm), '1', current_action, '', '',
             datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

    except Exception as error:
        Log([str(project_nm), str(report_nm), '0', current_action, str(error), traceback.format_exc(),
             datetime.now().strftime('%Y-%m-%d %H:%M:%S')])


def SharedFolder(project_nm, source_nm,
                 orig_folder_path, orig_file_name, dest_folder_path, dest_file_name=None,
                 arch_folder_path=None, add_time_to_archive=False,
                 orig_filename_date_offset=None, orig_filename_month_offset=None, add_orig_file_name_col=False,
                 csv_encoding='utf_8',
                 excel_sheet=0, skip_header_rows=0, skip_footer_rows=0, data_types=None, date_columns=None,
                 power_query_name=None, power_query_sheet=1, power_query_table=1,
                 power_query_datetime_columns_formats=None,
                 zip_short_file_names_str=None, add_orig_zip_file_name_file_names_str=None,
                 delete_orig_file=False, power_query_source_path_to_delete=None):
    # DESCRIPTION:
    # Takes file from specified origin folder and saves it in specified destination folder;
    # the supported file formats are: csv, xls, xlsx, xlsb, ods, zip;
    # if the file is CSV - column with the origin file name could be added;
    # if the file is Excel - converts it to CSV and column with the origin Excel file name could be added to the CSV;
    # if the file is Excel with Power Query - first the Power Query is refreshed and then optionally the source folder
    # of the Power Query is cleared from files;
    # if the file is ZIP - its contents are extracted and optionally the file name(s) are renamed based on
    # search string(s) and the origin ZIP file name is added as column for the specified CSV files extracted;
    # optionally deletes the origin file;
    #
    # PARAMETERS:
    # project_nm: name of the project for which the extraction is performed
    # source_nm: name of the data source related with the file being extracted
    # orig_folder_path: the folder path of the origin file
    # orig_file_name: the file name of the origin file; if date-formatting characters are present in the name,
    #                 it could be dynamic and works in combination with 'orig_filename_date_offset' or
    #                 'orig_filename_month_offset' depending on whether you want to have previous data or previous month
    #                 in the file name (one of these parameters should be set)
    # dest_folder_path: the folder path to which the origin file will be saved
    # dest_file_name (optional): the file name after it is saved in the destination folder;
    #                            if not set, the 'orig_file_name' will be used
    # arch_folder_path (optional): the folder path where the origin file will be archived;
    #                              if not set, the attachment won't be archived
    # add_time_to_archive (optional): if there is 'arch_folder_path' specified, shows whether the current date in
    #                                 'YYYY_MM_DD' format will be added at the end of the archive file name;
    #                                 the default value is False (no date will be added to the archive file name)
    # orig_filename_date_offset (optional): in case there are date-formatting characters in 'orig_file_name' - indicates
    #                                       how many days in the past will be the date in the 'orig_file_name';
    #                                       if you want today's date - use 0 (zero), for future dates use positive
    #                                       numbers; if not set and 'orig_filename_month_offset' is also not set,
    #                                       no date will be taken for 'orig_file_name'
    # orig_filename_month_offset (optional): in case there are date-formatting characters in 'orig_file_name' -
    #                                        indicates how many months in the past will be the date in the
    #                                        'orig_file_name'; if you want today's date - use 0 (zero),
    #                                        for future dates use positive numbers; if not set and
    #                                        'orig_filename_date_offset' is also not set, no date will be taken
    #                                        for 'orig_file_name';
    #                                        with precedence over 'orig_filename_date_offset'
    # add_orig_file_name_col (optional): indicates whether to add column with the origin file name (CSV or Excel)
    #                                    to the CSV file in the destination folder;
    #                                    the default is False (no column will be added)
    # csv_encoding (optional): in case 'add_orig_file_name_col' is True - indicates the encoding of the CSV file;
    #                          if not set, UTF-8 will be assumed
    # excel_sheet (optional): in case the origin file is Excel - indicates the sheet in the Excel file from which
    #                         the data will be taken; if not set, the first sheet will be read
    # skip_header_rows (optional): in case the origin file is Excel - indicates how many rows at the top of the Excel
    #                              sheet to be skipped when reading the file; if not set, no rows will be skipped
    # skip_footer_rows (optional): in case the origin file is Excel - indicates how many rows at the bottom of the Excel
    #                             sheet to be skipped when reading the file; if not set, no rows will be skipped
    # data_types (optional): in case the origin file is Excel - list with column names and data types with which
    #                        explicitly to set the data types in the pandas dataframe;
    #                        the following pattern should be followed (without spaces):
    #                        'column 1':'data type 1';'column 2':'data type 2';
    #                        if not set, the data types will be decided automatically
    # date_columns (optional): in case the origin file is Excel - list with column names separated by ';' indicating
    #                          the columns with dates which need to be converted after the Excel file is read;
    #                          if not set, no columns will be converted
    # power_query_name (optional): in case the origin file is Excel with Power Query - the name of the Power Query
    #                              connection to be refreshed;
    #                              if not set, no Power Query refresh will be performed
    # power_query_sheet (optional): in case the origin file is Excel with Power Query - the name or the index of the
    #                               sheet where the Power Query is refreshed;
    #                               if not set, it will take the first sheet
    # power_query_table (optional): in case the origin file is Excel with Power Query - the name or the index of the
    #                               table where the Power Query is refreshed;
    #                               if not set, it will take the first table in the chosen 'power_query_sheet'
    # power_query_datetime_columns_formats (optional): in case the origin file is Excel with Power Query -
    #                              string indicating the datetime columns which need to be formatted after the
    #                              Power Query is refreshed;
    #                              the following pattern should be followed:
    #                              'datetime column name 1'>'format string 1';'datetime column name 2'>'format string 2'
    #                              works in combination with 'power_query_sheet' and 'power_query_table';
    #                              if not set, no formatting will be done
    # zip_short_file_names_str (optional): in case the origin file is ZIP - list with keywords to search in the files
    #                                      of the zip file, separated by ';'; if any of the files in the zip contains
    #                                      the key word it will be renamed to that key word;
    #                                      if not set, no such search and renaming will be performed
    # add_orig_zip_file_name_file_names_str (optional): in case the origin file is ZIP - list with the file names of the
    #                                                   extracted csv files for which you need column with the zip file
    #                                                   name (including the extension and after their eventual renaming
    #                                                   (see previous parameter)) and optionally their encoding
    #                                                   (if encoding is not mentioned, UTF-8 will be assumed);
    #                                                   the following pattern should be followed:
    #                                                   'filename 1':'encoding 1';'filename 2':'encoding 2'
    # delete_orig_file (optional): in case the origin file is CSV or Excel - indicates whether to delete the origin
    #                              file; the default is False (origin file won't be deleted)
    # power_query_source_path_to_delete (optional): in case the origin file is Excel with Power Query - the source
    #                                               folder path of the Power Query which needs to be cleared;
    #                                               if not set, no files from the Power Query source folder will be
    #                                               deleted

    current_action = 'create origin file name with date offset'
    try:
        if orig_filename_date_offset is not None:
            orig_file_name = (datetime.today() - dateutil.relativedelta.relativedelta(days=orig_filename_date_offset)).strftime(orig_file_name)
        elif orig_filename_month_offset is not None:
            orig_file_name = (datetime.today() - dateutil.relativedelta.relativedelta(months=orig_filename_month_offset)).strftime(orig_file_name)

        current_action = 'create origin and destination paths'
        orig_file_path = str(os.path.join(orig_folder_path, orig_file_name))
        if dest_file_name is None:
            dest_file_name = orig_file_name
        dest_file_path = str(os.path.join(dest_folder_path, dest_file_name))

        current_action = 'save origin file in the archive'
        if arch_folder_path is not None:
            CopyFile(orig_file_path, orig_file_name, arch_folder_path, add_time_to_archive)

        current_action = 'check origin file extension'
        if (orig_file_name.endswith('.csv')
                or orig_file_name.endswith('csv')):
            current_action = 'copy CSV file'
            CopyCSV(orig_file_path, orig_file_name, dest_file_path,
                    add_orig_file_name_col, csv_encoding, delete_orig_file)
            current_action = 'CSV file saved in destination folder'
            Log([str(project_nm), str(source_nm), '1', current_action, '', '',
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

        elif (orig_file_name.endswith('.xlsx')
              or orig_file_name.endswith('.xls')
              or orig_file_name.endswith('.xlsb')
              or orig_file_name.endswith('.ods')):
            current_action = 'refresh Power Query in Excel file'
            if power_query_name is not None:
                power_query_errors = RefreshPowerQuery(orig_file_path, power_query_name, power_query_sheet,
                                                       power_query_table, power_query_datetime_columns_formats)
                if power_query_errors[0] is not None:
                    Log([str(project_nm), str(source_nm), '0', current_action, power_query_errors[0],
                         power_query_errors[1], datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                    return
                if power_query_source_path_to_delete is not None:
                    ClearFolder(power_query_source_path_to_delete)
            current_action = 'convert Excel to CSV'
            ExcelToCSV(orig_file_path, orig_file_name, dest_file_path, add_orig_file_name_col,
                       excel_sheet, skip_header_rows, skip_footer_rows,
                       data_types, date_columns, delete_orig_file)
            current_action = 'Excel file saved as CSV in destination folder'
            Log([str(project_nm), str(source_nm), '1', current_action, '', '',
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

        elif orig_file_name.endswith('.zip'):
            try:
                current_action = 'extract ZIP file'
                ExtractZIP(orig_file_path, orig_file_name, dest_folder_path,
                           zip_short_file_names_str, add_orig_zip_file_name_file_names_str)
                current_action = 'ZIP file extracted in destination folder'
                Log([str(project_nm), str(source_nm), '1', current_action, '', '',
                     datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
            except Exception as error:
                Log([str(project_nm), str(source_nm), '0', current_action, str(error), traceback.format_exc(),
                     datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                return

        else:
            current_action = 'file format is not supported'
            Log([str(project_nm), str(source_nm), '0', current_action, '', '',
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

    except Exception as error:
        Log([str(project_nm), str(source_nm), '0', current_action, str(error), traceback.format_exc(),
             datetime.now().strftime('%Y-%m-%d %H:%M:%S')])


def Email(project_nm, source_nm, subject_contains_str, file_ext, dest_folder_path,
          email_folder_path=None, file_name_contains_str=None, save_as_file_name=None,
          target_time_sign=None, target_time_str=None,
          arch_folder_path=None, add_time_to_archive=False, add_time_to_source=False,
          excel_to_csv=True, add_orig_excel_file_name_col=False, excel_sheet=0, skip_header_rows=0, skip_footer_rows=0,
          data_types=None, date_columns=None,
          archive_email=True, archive_email_folder='Processed Emails'):
    # DESCRIPTION:
    # Gets email attachments from the predefined mailbox;
    # Searches by a string in the subject name and specific extension of the file
    # (and optionally by string in the file name and/or the time the email arrived);
    # Saves the attachment in specific folder with specific name, if mentioned.
    # If the attachment is Excel - converts it to CSV file;
    # If the attachment is zip file - extracts its contents the destination folder;
    # Moves the processed emails in the predefined folder in the mailbox;
    # optionally archives the attachment in specific folder and adds timestamp of the email to the name
    #
    # PARAMETERS:
    # project_nm: name of the project for which the extraction is performed
    # source_nm: name of the data source related with the file being extracted
    # subject_contains_str: key word which is searched in the emails to decide whether to process them further
    # file_ext: file extension which is searched in the email attachments to decide whether to process them further
    # dest_folder_path: the folder path to which the attachment is saved
    # email_folder_path (optional): the folder path in the mailbox where the emails to be searched for are located;
    #                               '>' is used as separator between the folders;
    #                               if not set, the Inbox folder will be searched
    # file_name_contains_str (optional): key word which is searched in the email attachment to decide whether to
    #                                    pick it; if not set, no search in the email attachment name is performed
    # save_as_file_name (optional): indicates how to change the file name of the email attachment when saving it to the
    #                               destination folder; if not set, the email attachment name won't be changed
    # target_time_sign (optional): specifies before or after which time should be the email sent time in order for the
    #                              attachment to be picked; works in combination with 'target_time_str';
    #                              if not set, there won't be a condition even if there may be 'target_time_str' set
    # target_time_str (optional): specifies the time before or after which the sent time of the email should be in order
    #                             for the attachment to be picked; works in combination with 'target_time_sign';
    #                             if not set, there won't be a condition even if there may be 'target_time_sign' set
    # arch_folder_path (optional): the folder path where the attachment will be archived;
    #                              if not set, the attachment won't be archived
    # add_time_to_archive (optional): if there is 'arch_folder_path' specified, shows whether the timestamp of the email
    #                                 will be added at the end of the archive file name;
    #                                 the default value is False (no timestamp will be added to the archive file name)
    # add_time_to_source (optional): shows whether the timestamp of the email will be added at the end of the file name
    #                                saved in the 'dest_folder_path';
    # excel_to_csv (optional): indicates whether to convert to CSV in case the saved attachment is Excel file;
    #                          the default is True (the convertion to CSV will be performed)
    # add_orig_excel_file_name_col (optional): in case 'excel_to_csv' is True - indicates whether to add column with
    #                                          the Excel file name to the CSV file;
    #                                          the default is False (no column will be added)
    # excel_sheet (optional): in case 'excel_to_csv' is True - indicates the sheet in the Excel file from which
    #                         the data will be taken; if not set, the first sheet will be read
    # skip_header_rows (optional): in case 'excel_to_csv' is True - indicates how many rows at the top of the Excel
    #                              sheet to be skipped when reading the file; if not set, no rows will be skipped
    # skip_footer_rows (optional): in case 'excel_to_csv' is True - indicates how many rows at the bottom of the Excel
    #                              sheet to be skipped when reading the file; if not set, no rows will be skipped
    # data_types (optional): in case 'excel_to_csv' is True - list with column names and data types with which
    #                        explicitly to set the data types in the pandas dataframe;
    #                        the following pattern should be followed (without spaces):
    #                        'column 1':'data type 1';'column 2':'data type 2';
    #                        if not set, the data types will be decided automatically
    # date_columns (optional): in case 'excel_to_csv' is True - list with column names separated by ';' indicating
    #                          the columns with dates which need to be converted after the Excel file is read;
    #                          if not set, no columns will be converted
    # archive_email (optional): moves the email(s) containing the key word in 'subject_contains_str' to the predefined
    #                           archive folder ('archive_email_folder') in the mailbox;
    #                           the default is True (the email will be moved to the 'archive_email_folder')
    # archive_email_folder (optional): the archive folder of the mailbox;
    #                                  if not set specifically it's 'Processed Emails'

    email_data_extracted = False
    current_action = 'create Outlook Folder and Items objects'
    try:
        username = str.lower(os.getlogin())
        outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        # if username == 'service_279_test_rep':
        #     inbox_folder = outlook.Folders('crt_uk@foundever.com').Folders('Inbox')
        # else:
        #     inbox_folder = outlook.Folders('Mailbox CRT UK').Folders('Inbox')
        inbox_folder = outlook.Folders('Mailbox CRT UK').Folders('Inbox')
        outlook_folder = inbox_folder
        if email_folder_path is not None:
            email_folder_path = email_folder_path.split('>')
            for email_folder in email_folder_path:
                email_folder = email_folder.strip()
                outlook_folder = outlook_folder.Folders(email_folder)
        outlook_emails = outlook_folder.Items
        # sort from old to new
        outlook_emails.Sort('ReceivedTime', False)

        current_action = 'loop through Outlook folder'
        for outlook_email in outlook_emails:
            current_action = 'check if the email contains the key word'
            if str.lower(subject_contains_str) in str.lower(outlook_email.Subject):
                outlook_email.Unread = False
                try:
                    email_timestamp = outlook_email.ReceivedTime.strftime('%Y_%m_%d_%H_%M_%S')
                except:
                    email_timestamp = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')

                current_action = 'check target time condition'
                email_time = datetime.strptime(email_timestamp, "%Y_%m_%d_%H_%M_%S").time()
                target_time = None
                if target_time_str is not None:
                    target_time = datetime.strptime(target_time_str, "%H:%M").time()
                if (target_time_sign is None or target_time_str is None
                        or (target_time_sign == '<' and email_time < target_time)
                        or (target_time_sign == '<=' and email_time <= target_time)
                        or (target_time_sign == '>' and email_time > target_time)
                        or (target_time_sign == '>=' and email_time >= target_time)):

                    current_action = 'loop through the email attachments'
                    for attachment in outlook_email.Attachments:
                        current_action = 'check if the file extension and file name are correct'
                        if (str.lower(attachment.Filename).endswith(str.lower(file_ext))
                                and (file_name_contains_str is None
                                     or str.lower(file_name_contains_str) in str.lower(attachment.Filename))):
                            current_action = 'save attachment in the archive'
                            if arch_folder_path is not None:
                                SaveAttachment(attachment, arch_folder_path,
                                               add_suffix_to_file_name=add_time_to_archive,
                                               suffix=email_timestamp)

                            current_action = 'save attachment in destination folder'
                            dest_file_path = SaveAttachment(attachment, dest_folder_path, save_as_file_name,
                                                            add_suffix_to_file_name=add_time_to_source,
                                                            suffix=email_timestamp)

                            if (dest_file_path.endswith('.xlsx')
                               or dest_file_path.endswith('.xls')
                               or dest_file_path.endswith('.xlsb')
                               or dest_file_path.endswith('.ods')) and excel_to_csv:
                                current_action = 'convert Excel attachment to CSV'
                                ExcelToCSV(dest_file_path, attachment.Filename, dest_file_path,
                                           add_orig_excel_file_name_col, excel_sheet,
                                           skip_header_rows, skip_footer_rows,
                                           data_types, date_columns, delete_orig_file=True)

                            if dest_file_path.endswith('.zip'):
                                current_action = 'extract file(s) from ZIP attachment'
                                ExtractZIP(dest_file_path, attachment.Filename, dest_folder_path, delete_orig_file=True)

                            email_data_extracted = True

        current_action = 'archive email'
        if archive_email:
            MoveEmails(outlook_emails, subject_contains_str, inbox_folder.Folders(archive_email_folder))

        if email_data_extracted:
            current_action = 'email data extracted'
            Log([str(project_nm), str(source_nm), '1', current_action, '', '',
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        else:
            current_action = 'email with data not found'
            Log([str(project_nm), str(source_nm), '0', current_action, '', '',
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

    except Exception as error:
        Log([str(project_nm), str(source_nm), '0', current_action, str(error), traceback.format_exc(),
             datetime.now().strftime('%Y-%m-%d %H:%M:%S')])


def Log(values):
    # DESCRIPTION:
    # Inserts row with the result of the script execution to the predefined SQL Server log table
    #
    # PARAMETERS:
    # values: array with the values needed by the SQL statement inserting in the log table

    server_nm = 'EU999K16SSCDB01'
    database_nm = 'gen_bg_attr'
    table_nm = 't_log'
    conn = pyodbc.connect('DRIVER={SQL Server}' +
                          ';SERVER=' + server_nm +
                          ';DATABASE=' + database_nm +
                          ';Trusted_Connection=yes')
    cursor = conn.cursor()
    sql_command = 'INSERT INTO ' + table_nm + ' VALUES (?,?,?,?,?,?,?)'
    cursor.execute(sql_command, values)
    cursor.commit()
    conn.close()


def CopyCSV(orig_file_path, orig_file_name, dest_file_path,
            add_orig_file_name_col=False, csv_encoding='utf_8', delete_orig_file=False):
    # DESCRIPTION:
    # Copies file from one location to another with possibility to change the file name;
    # optionally adds the origin file name as separate column (for CSV files) and/or delete the origin file
    #
    # PARAMETERS:
    # orig_file_path: file path of the origin file
    # orig_file_name: file name of the origin file
    # dest_file_path: file path of the destination file
    # add_orig_file_name_col (optional): indicates whether to add column with the origin file name to the CSV file;
    #                                    the default is False (no column will be added)
    # csv_encoding (optional): in case 'add_orig_file_name_col' is True - indicates the encoding of the CSV file;
    #                          if not set, UTF-8 will be assumed
    #
    # delete_orig_file (optional): indicates whether to delete the origin file;
    #                              the default is False (origin file won't be deleted)

    shutil.copy(orig_file_path, dest_file_path)
    if add_orig_file_name_col:
        csv_df = pd.read_csv(dest_file_path, encoding=csv_encoding)
        csv_df['File Name'] = orig_file_name
        csv_df.to_csv(dest_file_path, index=False)
    if delete_orig_file:
        time.sleep(2)
        os.remove(orig_file_path)


def ExcelToCSV(orig_file_path, orig_file_name, dest_file_path, add_orig_file_name_col=False,
               excel_sheet=0, skip_header_rows=0, skip_footer_rows=0,
               data_types=None, date_columns=None, delete_orig_file=False):
    # DESCRIPTION:
    # Converts Excel file to CSV file (reads Excel file as pandas dataframe and saves the dataframe as CSV file);
    # optionally can add column with the Excel file name to the CSV file and/or delete the Excel file
    #
    # PARAMETERS:
    # orig_file_path: the file path of the Excel file which will be converted
    # orig_file_name: the file name of the Excel file (with the extension)
    # dest_file_path: the file path of the CSV file where the pandas dataframe will be saved
    # add_orig_file_name_col (optional): indicates whether to add column with the Excel file name to the CSV file;
    #                                    the default is False (no column will be added)
    # excel_sheet (optional): the sheet in the Excel file from which the data will be taken;
    #                         if not set, the first sheet will be read
    # skip_header_rows (optional): indicates how many rows at the top of the Excel sheet to be skipped
    #                              when reading the file; if not set, no rows will be skipped
    # skip_footer_rows (optional): indicates how many rows at the bottom of the Excel sheet to be skipped
    #                              when reading the file; if not set, no rows will be skipped
    # data_types (optional): list with column names and data types with which explicitly to set the data types in the
    #                        pandas dataframe; the following pattern should be followed (without spaces):
    #                        'column 1':'data type 1';'column 2':'data type 2';
    #                        if not set, the data types will be decided automatically
    # date_columns (optional): list with column names separated by ';' indicating the columns with dates which need
    #                          to be converted after the Excel file is read; if not set, no columns will be converted
    # delete_orig_file (optional): indicates whether to delete the Excel file;
    #                              the default is False (Excel file won't be deleted)

    if data_types is not None:
        data_types = dict([row.split(':') for row in data_types.split(';')])
    excel_df = pd.read_excel(orig_file_path, sheet_name=excel_sheet,
                             skiprows=skip_header_rows, skipfooter=skip_footer_rows,
                             dtype=data_types, keep_default_na=False)
    if date_columns is not None:
        date_columns = date_columns.split(';')
        if orig_file_name.endswith('.xlsb'):
            for date_column in date_columns:
                date_column = date_column.strip()
                excel_df[date_column] = excel_df[date_column].apply(lambda x: pyxlsb.convert_date(x))
        else:
            for date_column in date_columns:
                date_column = date_column.strip()
                excel_df[date_column] = pd.to_datetime(pd.to_numeric(excel_df[date_column]),
                                                       unit='D', origin='1899-12-30')
    if add_orig_file_name_col:
        excel_df['File Name'] = orig_file_name
    if not dest_file_path.endswith('.csv'):
        dest_file_path = ChangeFilePathExtension(dest_file_path, 'csv')
    excel_df.to_csv(dest_file_path, index=False)
    if delete_orig_file:
        time.sleep(2)
        os.remove(orig_file_path)


def ExtractZIP(orig_file_path, orig_file_name, dest_folder_path,
               zip_short_file_names_str=None, add_orig_zip_file_name_file_names_str=None, delete_orig_file=False):
    # DESCRIPTION:
    # Extracts file(s) from the specified zip file;
    # can search the extracted file name(s) for specified text and rename the file name to it if it contains it;
    # can add a column with the original zip file name if the extracted file(s) is CSV;
    # optionally deletes the zip file
    #
    # PARAMETERS:
    # orig_file_path: the file path of the zip file
    # orig_file_name: the zip file name
    # dest_folder_path: the folder path where the zip will is extracted
    # zip_short_file_names_str (optional): list with keywords to search in the files of the zip file, separated by ';';
    #                                      if any of the files in the zip contains the key word it will be renamed to
    #                                      that key word; if not set, no such search and renaming will be performed
    # add_orig_zip_file_name_file_names_str (optional): list with the file names of the extracted csv files for which
    #                                                   you need column with the zip file name (including the extension
    #                                                   and after their eventual renaming (see previous parameter))
    #                                                   and optionally their encoding (if encoding is not mentioned,
    #                                                   UTF-8 will be assumed);
    #                                                   the following pattern should be followed:
    #                                                   'filename 1':'encoding 1';'filename 2':'encoding 2'
    #                                                   works in combination with 'orig_file_name' which carries
    #                                                   information about the zip file name
    # delete_orig_file (optional): indicates whether to delete the zip file;
    #                              the default is False (zip file won't be deleted)

    zip_file = zipfile.ZipFile(orig_file_path, 'r')
    zip_files = zip_file.namelist()
    zip_short_file_names = None
    if zip_short_file_names_str is not None:
        zip_short_file_names = zip_short_file_names_str.split(';')
    add_orig_zip_file_name_file_names = None
    if add_orig_zip_file_name_file_names_str is not None:
        add_orig_zip_file_name_file_names = [row.split(':') for row in add_orig_zip_file_name_file_names_str.split(';')]
    for zip_file_name in zip_files:
        zip_file.extract(zip_file_name, dest_folder_path)
        extracted_file_name = zip_file_name
        if zip_short_file_names_str is not None:
            for zip_short_file_name in zip_short_file_names:
                zip_short_file_name = zip_short_file_name.strip()
                if zip_short_file_name in zip_file_name and zip_short_file_name != '':
                    extracted_file_name = RenameFullFileName(zip_file_name, zip_short_file_name)
                    old_file_path = str(os.path.join(dest_folder_path, zip_file_name))
                    new_file_path = str(os.path.join(dest_folder_path, extracted_file_name))
                    if os.path.exists(new_file_path):
                        os.remove(new_file_path)
                    os.rename(old_file_path, new_file_path)
        if add_orig_zip_file_name_file_names_str is not None and extracted_file_name.endswith('.csv'):
            for add_orig_zip_file_name_file_name in add_orig_zip_file_name_file_names:
                if add_orig_zip_file_name_file_name[0].strip() == extracted_file_name:
                    if len(add_orig_zip_file_name_file_name) == 1:
                        extracted_file_csv_encoding = 'utf-8'
                    else:
                        extracted_file_csv_encoding = add_orig_zip_file_name_file_name[1].strip()
                    csv_file_path = str(os.path.join(dest_folder_path, extracted_file_name))
                    csv_df = pd.read_csv(csv_file_path, encoding=extracted_file_csv_encoding)
                    csv_df['File Name'] = orig_file_name
                    csv_df.to_csv(csv_file_path, index=False)
    time.sleep(2)
    zip_file.close()
    if delete_orig_file:
        os.remove(orig_file_path)


def CopyFile(orig_file_path, orig_file_name, dest_folder_path, add_date_to_file_name=False):
    # DESCRIPTION:
    # Copies a file to another shared folder;
    # optionally can add suffix to the file name (in the destination) with the current date
    #
    # PARAMETERS:
    # orig_file_path: the file path of the file which will be copied to another location
    # orig_file_name: the name of the file (with the extension) which will be copied to another location
    # dest_folder_path: the folder path where the file will be copied
    # add_date_to_file_name (optional): indicates whether to add date at the end of the file name
    #                                   (in 'YYYY_MM_DD' format) when saving it to the destination folder;
    #                                   the default is False (no date will be added to the file name)

    dest_file_name = orig_file_name
    if add_date_to_file_name:
        dest_file_name = AddSuffixToFileName(dest_file_name, datetime.now().strftime('%Y_%m_%d'))
    dest_file_path = str(os.path.join(dest_folder_path, dest_file_name))
    shutil.copy(orig_file_path, dest_file_path)


def MoveEmails(outlook_emails, move_email_containing_str, dest_email_folder):
    # DESCRIPTION:
    # Moves emails containing specific key word in their subjects to another mailbox folder
    #
    # PARAMETERS:
    # outlook_emails: collection of Outlook emails whose subjects will be searched for the key word
    #                 and eventually moved to another folder if they contain it
    # move_email_containing_str: the key word by which the subjects of the emails will be searched to determine
    #                            whether they should be moved to another folder
    # dest_email_folder: the Outlook folder object to which the emails containing the key word
    #                    in its subject will be moved

    i = outlook_emails.Count
    while i > 0:
        if str.lower(move_email_containing_str) in str.lower(outlook_emails(i).Subject):
            outlook_emails(i).Move(dest_email_folder)
        i = i - 1


def SaveAttachment(attachment, dest_folder_path, save_as_file_name=None, add_suffix_to_file_name=False, suffix=None):
    # DESCRIPTION:
    # Saves the specified email attachment in a shared folder and returns the file path;
    # optionally can change the file name and/or add suffix at the end of the file name
    #
    # PARAMETERS:
    # attachment: Outlook email attachment object which will be saved
    # dest_folder_path: the path of the folder where the attachment is going to be saved
    # save_as_file_name (optional): used to change the name of the attachment; if not set the name won't be changed
    # add_suffix_to_file_name (optional): used to add suffix at the end of the file name (can be used in combination
    #                                     with 'save_as_file_name' which will first change the file name and then the
    #                                     suffix will be added); the default is False (no suffix will be added)
    # suffix (optional): the string which will be added in case 'add_time_to_file_name' is set to True;
    #                    if not set, no suffix will be added

    dest_file_name = attachment.Filename
    if save_as_file_name is not None:
        dest_file_name = RenameFullFileName(dest_file_name, save_as_file_name)
    if add_suffix_to_file_name:
        dest_file_name = AddSuffixToFileName(dest_file_name, suffix)
    dest_file_path = os.path.join(dest_folder_path, dest_file_name)
    attachment.SaveAsFile(dest_file_path)
    return dest_file_path


def RefreshPowerQuery(excel_filepath, power_query_name, power_query_sheet=1, power_query_table=1,
                      power_query_datetime_columns_formats=None, save_after_close=True):
    # DESCRIPTION:
    # Refreshes one Power Query connection in the specified Excel file and returns eventual errors
    # optionally changes the datetime format of chosen columns
    #
    # PARAMETERS:
    # excel_filepath: the file path of the Excel file containing the Power Query
    # power_query_name: the name of the Power Query connection to be refreshed
    # power_query_sheet (optional): the name or the index of the sheet where the Power Query is refreshed;
    #                               if not set, it will take the first sheet
    # power_query_table (optional): the name or the index of the table where the Power Query is refreshed;
    #                               if not set, it will take the first table in the chosen 'power_query_sheet'
    # power_query_datetime_columns_formats (optional): string indicating the datetime columns which
    #                              need to be formatted; the following pattern should be followed:
    #                              'datetime column name 1'>'format string 1';'datetime column name 2'>'format string 2'
    #                              works in combination with 'power_query_sheet' and 'power_query_table';
    #                              if not set, no formatting will be done
    # save_after_close (optional): indicates whether to save the Excel file after the Power Query connection is
    #                              refreshed; the default is True (saves the Excel file)

    errors = (None, None)
    excel_app = win32com.client.DispatchEx("Excel.Application")
    excel_wb = excel_app.Workbooks.Open(excel_filepath)
    try:
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_wb.Connections('Query - ' + power_query_name).Refresh()
        if power_query_datetime_columns_formats is not None:
            excel_tbl = excel_wb.Worksheets(power_query_sheet).ListObjects(power_query_table)
            power_query_datetime_columns_formats = [row.split('>') for row in power_query_datetime_columns_formats.split(';')]
            for power_query_datetime_columns_format in power_query_datetime_columns_formats:
                datetime_column = power_query_datetime_columns_format[0].strip()
                datetime_format = power_query_datetime_columns_format[1].strip()
                excel_tbl.ListColumns(datetime_column).Range.NumberFormat = datetime_format
        excel_wb.Close(save_after_close)
    except Exception as error:
        excel_wb.Close(False)
        errors = (str(error), traceback.format_exc())
    excel_app.DisplayAlerts = True
    excel_app.Visible = True
    excel_app.Quit()
    return errors


def ClearFolder(folder_path):
    # DESCRIPTION:
    # Clears every file in the specified folder
    #
    # PARAMETERS:
    # folder_path: the path of the folder to be cleared

    files = os.listdir(folder_path)
    time.sleep(2)
    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        os.remove(file_path)


def AddSuffixToFileName(filename_full, suffix):
    # DESCRIPTION:
    # Returns full filename (with the extension) with added suffix at the end of the filename
    #
    # PARAMETERS:
    # filename_full: the name of the file with the extension
    # suffix: the string to add at the end of the filename

    ext_position = filename_full.rfind('.')
    ext_length = len(filename_full) - ext_position
    filename = filename_full[:ext_position]
    ext = filename_full[-ext_length:]
    return filename + '_' + str(suffix) + ext


def RenameFullFileName(filename_full, new_filename):
    # DESCRIPTION:
    # Changes the filename and returns the full filename (with the extension)
    #
    # PARAMETERS:
    # filename_full: the name of the file with the extension
    # new_filename: the filename which will replace the old one (without the extension)

    ext_position = filename_full.rfind('.')
    ext_length = len(filename_full) - ext_position
    ext = filename_full[-ext_length:]
    return new_filename + ext


def ChangeFilePathExtension(file_path, new_ext):
    # DESCRIPTION:
    # Changes the extension of the full file path (ending with the extension) and returns the new file path
    #
    # PARAMETERS:
    # file_path: the file path with the extension
    # new_ext: the new extension which will replace the old one (without the leading dot)

    ext_position = file_path.rfind('.')
    return file_path[:ext_position + 1] + new_ext
