"""This is the ExcelReader module.

This module is download the xlsx from the given url,
then it will read the sheet with given name
then it will convert the row data into a list of dict except row 1,
the values in row one would be the keys for in each dict.
Then the result will be stored as a .json file in aws s3

"""

__version__ = '0.1'
__author__ = 'Jobi A J'


import os
import json
import requests
import sys
import time
import uuid
import boto3
from xlrd import open_workbook
from botocore.client import ClientError


class ExcelReaderScript():

    def __init__(self):
        self.default_url = "https://www.iso20022.org/sites/default/files/ISO10383_MIC/ISO10383_MIC.xls"
        self.default_sheet_name = "MICs List by CC"

    def run(self):
        # Downloading the file from the given url, saving it a file.
        # Reading the file and writing to a json file
        result_file = self.read_excel_file()
        # The created result file is pushing to amazone s3
        self.push_picture_to_s3(result_file)

    def download_file_from_url(self):
        url = self.default_url
        try:
            r = requests.get(url, allow_redirects=True)
            current_working_directory = os.getcwd()
            file_name = str(uuid.uuid4()) + 'test.xlsx'
            # Downloading the file to the working directory
            downloaded_file = os.path.join(current_working_directory,
                                           file_name)
            open(downloaded_file, 'wb').write(r.content)
        except requests.exceptions.RequestException as e:
            print("Error occured while trying to connect the given url %s" % e)
            sys.exit(1)
        except Exception as e:
            print("Error occured during \
                   downloading the file %s" % e)
            sys.exit(1)
        else:
            print("File successfully downloaded.")
            return downloaded_file

    def read_excel_file(self):
        excel_file_path = self.download_file_from_url()
        wb = open_workbook(excel_file_path)
        required_sheet = []
        sheets = wb.sheets()
        for sheet in sheets:
            if sheet.name == self.default_sheet_name:
                required_sheet.append(sheet)
        if len(required_sheet) == 0:
            print("Unable to find the required excel sheet tab")
            sys.exit(1)
        else:
            try:
                sheet = required_sheet[0]
                number_of_rows = sheet.nrows
                number_of_columns = sheet.ncols
                # Getting the keys of the each row (Header)
                keys = [sheet.cell(0, col_index).value
                        for col_index in range(number_of_columns)]

                dict_list = []
                for row_index in range(1, number_of_rows):
                    d = {keys[col_index]: sheet.cell(row_index,
                         col_index).value
                         for col_index in range(number_of_columns)}
                    dict_list.append(d)
                data = json.dumps(dict_list, indent=4, sort_keys=True)
                current_working_directory = os.getcwd()
                file_name = ('processed_data' +
                             '_' +
                             str(uuid.uuid4()) +
                             '.json')
                # Storing the data as json file in current directory
                file = open(file_name, 'w')
                file.write(data)
                file.close()
            except Exception as e:
                print("Exception occured during reading the excel file \
                    or creating the required output %s" % e)
                sys.exit(1)
            else:
                # Delete the xlsx file after processing
                os.remove(excel_file_path)
                return file_name

    def push_picture_to_s3(self, filename):
        s3 = boto3.resource('s3')
        # Provide your bucket name here
        my_bucket_name = "yourbucketname"
        my_bucket = s3.Bucket(my_bucket_name)
        try:
            s3.meta.client.head_bucket(Bucket=my_bucket.name)
        except ClientError:
            print("The bucket does not exist or you have no access.")
            sys.exit(1)
        else:
            timestr = time.strftime("%Y%m%d-%H%M%S")
            file_name_for_upload = (self.default_sheet_name +
                                    '_' +
                                    timestr +
                                    '.json')
            s3.Object(my_bucket_name,
                      file_name_for_upload).put(Body=open(filename, 'rb'))
        finally:
            print("Successfully uploaded to amazon s3")


if __name__ == "__main__":
    ExcelReader = ExcelReaderScript()
    ExcelReader.run()
