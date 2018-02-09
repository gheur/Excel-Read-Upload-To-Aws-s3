# Excel-Read-Upload-To-Aws-s3
Download excel file, read and create a json using the file data and store the output.

Steps: 
1. Download the xlsx from - https://www.iso20022.org/sites/default/files/ISO10383_MIC/ISO10383_MIC.xls
2. Store the xlsx
3. Read the tab titled "MICs List by CC"
4. Create a list of dict containing all rows (except row 1). The values in row 1 would be the keys for in each dict.
5. Store the list from step 4 as a .json file in aws s3

Prerequisites:
1. Packages
    1. os
    2. json
    3. requests
    4. sys
    5. time
    6. uuid
    7. boto3
    8. xlrd
    9. botocore
2. Create the amazone aws credential file in your local system. By default, its location is at ~/.aws/credentials (http://boto3.readthedocs.io/en/latest/guide/quickstart.html)
