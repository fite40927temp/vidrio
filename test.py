import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime

### SETUP ###
# Create exception file

# Copy mapping file to mapping subfolder
mapping_filepath = '../TestProject/Cash_Rec_Mapping.xlsx'
os.mkdir('Mapping')
shutil.copy(mapping_filepath, './Mapping')
mapping_filepath = './Mapping/Cash_Rec_Mapping.xlsx'

# Create subfolder for output
os.mkdir('Output')


### LOAD ###
# Ask for filepath
activity_file = input('Bank Activity filepath:\t')

# Read mapping and specified log into dataframes
activity = pd.read_excel(activity_file)
mapping = pd.read_excel(mapping_filepath)
activity = activity.fillna('')  # Replace NaN with empty string

# Exceptions logging
exceptions = pd.DataFrame()
any_exceptions = False


### MANAGE ###
# Bank Activity dataframe
activity['Bank Reference ID'] = activity['Reference Number']
activity['Post Date'] = activity['Cash Post Date']
activity['Value Date'] = activity['Cash Value Date']
activity['Amount'] = activity['Transaction Amount Local']
activity['Description'] = activity['Transaction Description 1'] + activity['Transaction Description 2'] + activity['Transaction Description 3'] + activity['Transaction Description 4'] + activity['Transaction Description 5'] + activity['Transaction Description 6'] + activity['Detailed Transaction Type Name'] + activity['Transaction Type']
activity['Bank Account'] = activity['Cash Account Number']
activity['Closing_Balance'] = activity['Closing Balance Local']
dt_string = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
activity['Filename'] = activity['Cash Account Number'].astype(str) + ' ' + dt_string + '.csv' #datetime.now()

# Mapping dataframe
mapping_bankid = pd.DataFrame(mapping['Bank Ref ID'])
mapping_idbalance = pd.concat([mapping['Bank Ref ID'], mapping['Starting_Balance']], axis=1)

for i in mapping_bankid['Bank Ref ID']:
    balance = mapping_idbalance[mapping_idbalance['Bank Ref ID'].eq(i)].iloc[0]['Starting_Balance']
    Output = activity[activity['Cash Account Number'].eq(i)]
    mm = Output['Description'][Output['Description'].str.contains('STIF')]
    Output = Output.drop(mm.index)
    write_file = pd.concat([Output['Bank Reference ID'],Output['Post Date'],Output['Value Date'],Output['Amount'],Output['Description'],Output['Bank Account'],Output['Closing_Balance']],axis=1)
    write_file['Bank Reference ID'] = write_file['Bank Reference ID'].replace('',np.nan)
    write_file = write_file[write_file['Bank Reference ID'].notna()]
    print(write_file)
    # if(write_file.empty):
    #     print(f'has no activity.')
    # else:
    #     print(write_file)