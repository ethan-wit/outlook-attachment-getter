import win32com.client
import pandas as pd
import zipfile
import datetime as dt
import warnings


class outlook_attachment():
    '''Class representing an outlook email attachment getter'''

    def __init__(self):

        self.attachment_filepath = None
        self.df = None


    def get_today_interval(self):
        '''
        Gets datetime for beginning of day (e.g. 2021-07-15 00:00:00) and current time (e.g. 2021-07-15 11:07:13).
        Can be used as start_interval, end_interval parameters in get_attachment, if searching for attachment that was received between 12am today and current time.
        @return start, end: tuple of datetime start and end intervals
        '''

        start_date = dt.date.today()
        start = dt.datetime(year=start_date.year, month=start_date.month, day=start_date.day,)

        end = dt.datetime.today()

        return start, end


    def get_attachment(self, folderpath_list: list, email_subject_name: str, email_attachment_name: str, start_interval=None, end_interval=None, save_attachment_as=None):
        '''
        Gets attachment from Outlook email within time interval from specified folder. Writes attachment to disk if valid filepath is input; otherwise, returns attachment as a 'win32com.client.CDispatch' object.
        If multiple emails match the query criteria, the most recent email + attachment will be chosen.
        @param folderpath_list: sequential list of subfolders, from highest folder to lowest subfolder (e.g. ['level_1', 'level_2'])
        @param email_subject_name: subject of desired email
        @param email_attachment_name: name of attachment on desired email
        @param start_interval: (optional) datetime object declaring left bound of email received time interval. *The start interval is included* (>=). Needs to include year, month day, hour, minute, and second.
        @param end_interval: (optional) datetime object declaring right bound of email received time interval. *The end interval is included* (<=). Needs to include year, month day, hour, minute, and second.
        @param save_attachment_as: (optional) filepath (path, file name, and file type must be included) where the attachment will be saved (e.g. "r\\network\attachment.xlsx")
        @return attachment: email attachment returned (if save_attachment_as is None)
        '''

        #Type checking
        assert(isinstance(folderpath_list, list)), f'''folderpath_list must be of type list. You input an object of type {type(folderpath_list)}.'''
        assert(isinstance(email_subject_name, str)), f'''email_subject_name must be of type string. You input an object of type {type(email_subject_name)}.'''
        assert(isinstance(email_attachment_name, str)), f'''email_attachment_name must be of type string. You input an object of type {type(email_attachment_name)}.'''
        if save_attachment_as is not None:
            assert(isinstance(save_attachment_as, str)), f'''save_attachment_as must be of type string. You input an object of type {type(save_attachment_as)}.'''
            #Ensure filepath is given (not just desired file name)
            if not (("\\" in save_attachment_as) or ("/" in save_attachment_as)):
                raise Exception(f'''Parameter save_attachment_as requires a full path. The method will not default to the current working directory.''')

        #Ensure start is <= end
        if (start_interval is not None) and (end_interval is not None):
            if start_interval > end_interval:
                raise Exception(f'''Start of interval {start_interval.strftime("%m/%d/%Y %H:%M %p")} is more recent than end of interval {end_interval.strftime("%m/%d/%Y %H:%M %p")}. Start of interval must come before or at the same time as end of interval.''')

        print(f'''Start interval email query: {start_interval.strftime("%m/%d/%Y %I:%M %p")}\n\n''')
        print(f'''End interval email query: {end_interval.strftime("%m/%d/%Y %I:%M %p")}\n\n''')

        #Interface with Microsoft Outlook (using IDispatch-based COM object)
        Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        #Create Folder object for folder that contains desired email & attachment
        for i, folder in enumerate(folderpath_list):

            if i==0:
                Folder = Outlook.Folders[folder]
            else:
                Folder = Folder.Folders[folder]

        #Get emails from folder
        emails = Folder.Items

        #Restrict emails to after or equal to start interval
        if start_interval is not None:
            assert(isinstance(start_interval, dt.datetime)), f'''start_interval must be of type datetime.datetime. You input an object of type {type(start_interval)}.'''
            emails = emails.Restrict("[ReceivedTime] >= '" + start_interval.strftime('%m/%d/%Y %H:%M %p') + "'")

        #Restrict emails to before or equal to end interval
        if end_interval is not None:
            assert(isinstance(end_interval, dt.datetime)), f'''end_interval must be of type datetime.datetime. You input an object of type {type(end_interval)}.'''
            emails = emails.Restrict("[ReceivedTime] <= '" + end_interval.strftime('%m/%d/%Y %H:%M %p') + "'")

        #Sort emails, descending order
        emails.Sort("[ReceivedTime]", True)

        #Preset email as not found
        found_attachment = False
        #Get most recent email+attachment that meets criteria
        for email in emails:

            if email.Subject == email_subject_name:

                for attachment in email.Attachments:

                    if attachment.FileName == email_attachment_name:

                        found_attachment = True
                        if save_attachment_as is not None:
                            try:
                                attachment.SaveAsFile(save_attachment_as)
                                print(f'''Most recent email attachment that meets criteria has been saved as {save_attachment_as}\n\n''')
                                self.attachment_filepath = save_attachment_as
                                return None
                            except:
                                raise Exception(f'''Attachment was not saved. Please refer to Python error.''')
                        else:
                            try:
                                return attachment
                            except:
                                raise Exception(f'''Attachment was not returned. Please refer to Python error.''')

        if found_attachment is False:
            raise Exception(f'''Could not find email {email_subject_name} with attachment {email_attachment_name} within (un)specified time interval.''')


    def extract_zip_content(self, filepath, extract_file_name: str, save_unzipped_folder: str, exact: bool = False):
        '''
        Unzips zip file
        @param extract_file_name: file (file name, and file type must be included) to be extracted from zip
        @param save_unzipped_folder: folderpath (entire folderpath (do NOT include file name or type) must be included) where the unzipped attachment will be saved (e.g. "r\\user\folder")
        @param exact: If False, will extract file from zip that has extract_file_name in its name. If True, will only extract file from zip with exact string as extract_file_name.
        '''

        #type checking
        assert(isinstance(exact, bool)), f'''exact must be of type bool. You input an object of type {type(exact)}.'''
        assert(isinstance(extract_file_name, str)), f'''extract_file_name must be of type string. You input an object of type {type(extract_file_name)}.'''
        assert(isinstance(save_unzipped_folder, str)), f'''save_unzipped_folder must be of type string. You input an object of type {type(save_unzipped_folder)}.'''

        if not (zipfile.is_zipfile(filepath)):
            print(f'''{filepath} is not a zip file; therefore it cannot be unzipped.''')
        else:
            #Create zipfile object with positions zip
            zip_file = zipfile.ZipFile(filepath)

            #Save specified file within zip
            file_found = False
            for file_name in zip_file.namelist():

                if exact is True:

                    if extract_file_name == file_name:

                        zip_file.extract(file_name, save_unzipped_folder)
                        filepath = str(save_unzipped_folder) + "\\" + file_name
                        file_found = True
                        print(f'''Unzipped attachment saved as {filepath}''')

                elif exact is False:

                    if extract_file_name in file_name:

                        zip_file.extract(file_name, save_unzipped_folder)
                        filepath = str(save_unzipped_folder) + "\\" + file_name

                        file_found = True
                        print(f'''Unzipped attachment saved as {filepath}''')

            if file_found is False:
                print(f'''Could not find file {file_name} in zipfile.''')


    def set_df(self):
        '''Will read excel file to pandas dataframe'''

        #Used to rid warning: "Workbook contains no default style, apply openpyxl's default"
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            self.df = pd.read_excel(self.attachment_filepath, engine='openpyxl')


    def df_getter(self):

        return self.df


if __name__ == "__main__":

    '''
    1. Instantiate attachment getter object
    '''

    #Outlook attachment getter object
    getter = outlook_attachment()

    '''
    2. Declare search interval
    '''

    #Get 12am today start interval, current time end interval
    #start_interval, end_interval = getter.get_today_interval()

    start_interval = dt.datetime(year=2021, month=7, day=27, hour=1)
    end_interval = dt.datetime(year=2021, month=7, day=27, hour=12)

    '''
    3. Get attachment and save to disk
    '''

    #Get most recent attachment with matching parameters
    getter.get_attachment(folderpath_list=['level_1', 'level_2'],
                          email_subject_name='email subject',
                          email_attachment_name='attachment.zip',
                          start_interval=start_interval,
                          end_interval=end_interval,
                          save_attachment_as=r'C:\user\test\test_attachment.zip')

