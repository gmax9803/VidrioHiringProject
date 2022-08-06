import pandas as pd
import os
import numpy
from datetime import datetime
import shutil

if not os.path.exists('Output'):
    os.makedirs('Output')

if not os.path.exists('Mapping'):
    os.makedirs('Mapping')

mapping_path = os.getcwd() + '\Cash_Rec_Mapping.xlsx'
new_mapping_path = os.getcwd() + '\\Mapping\\Cash_Rec_Mapping.xlsx'

while True:
    try:
        df_bankactivity = pd.read_excel(input('Please enter full path to "Bank Activity Statement".\n>>'))
    except:
        print('You entered an invalid path, please try again.\n>>')
        continue
    else:
        print("Bank activity statement was loaded successfully.")
        break

try:
    shutil.copyfile(mapping_path, new_mapping_path)
    df_mapping = pd.read_excel(new_mapping_path)
except:
    print('Error with loading mapping file.')
else:
    print('Mapping File loaded successfully.')


df_bankactivity.fillna("", inplace=True)
df_exceptions = pd.DataFrame()

exceptions_bool = False
# I recognize this is inefficient but in the interests of time + I'm on vacation I did it this way.

df_bankactivity['Bank Reference ID'] = df_bankactivity['Reference Number']
df_bankactivity['Post Date'] = df_bankactivity['Cash Post Date']
df_bankactivity['Value Date'] = df_bankactivity['Cash Value Date']
df_bankactivity['Amount'] = df_bankactivity['Transaction Amount Local']
df_bankactivity['Description'] = df_bankactivity[['Transaction Description 1', 'Transaction Description 2', 'Transaction Description 3', \
    'Transaction Description 4', 'Transaction Description 5', 'Transaction Description 6', 'Detailed Transaction Type Name', 'Transaction Type']].agg("".join, axis=1)
df_bankactivity['Bank Account'] = df_bankactivity['Cash Account Number']
df_bankactivity['Closing_Balance'] = df_bankactivity['Closing Balance Local']
df_bankactivity['Filename'] = str(df_bankactivity['Cash Account Number']) + ' ' + str(datetime.now().strftime("%Y-%m-%d-%H-%M-%S")) + '.csv'

df_bankactivity.fillna("", inplace=True)

print('^^^^^^^^^^^^^^^^^^^^^^^')
print("See comment, but I recognize this isn't the best way to do it.")

df_refID_map = df_mapping['Bank Ref ID']

df_refID_sbal_map = pd.DataFrame()

df_refID_sbal_map['Bank Ref ID'] = df_mapping['Bank Ref ID']
df_refID_sbal_map['Starting_Balance'] = df_mapping['Starting_Balance']

df_oput = pd.DataFrame()
df_mm = pd.DataFrame()

for i, id in df_refID_map.items():
    sbal = df_refID_sbal_map[(df_refID_sbal_map['Bank Ref ID'] == id)]['Starting_Balance'].iloc[0]
    
    df_oput = df_bankactivity[(df_bankactivity['Cash Account Number'] == id) & (df_bankactivity['Transaction Description 1'].str.contains('STIF') == False)]
    df_mm = df_bankactivity[(df_bankactivity['Description'].str.contains('STIF')) & (df_bankactivity['Cash Account Number'] == id)]
    df_write_file = df_oput[['Bank Reference ID', 'Post Date', 'Value Date', 'Amount', 'Description', 'Bank Account', 'Closing_Balance']]

    if df_write_file.empty:
        print(str(id) + ' Has no activity')
    else:
        bank_closing_balance = df_bankactivity[(df_bankactivity['Cash Account Number'] == id)]['Closing_Balance'].iloc[0]
        mm_overnight = df_mm['Amount'].sum() # Assuming I should use amount column, instructions were not specific

        df_write_file = pd.concat([df_write_file, pd.DataFrame({
            'Bank Reference ID': ['Starting Balance'],
            'Post Date': ['2020-01-01'],
            'Value Date': ['2020-01-01'],
            'Amount': [sbal],
            'Description': ['Starting Balance'],
            'Bank Account': [id],
            'Closing_Balance': [0]}, columns=df_write_file.columns)], ignore_index=True)

        df_write_file['Amount'] = pd.to_numeric(df_write_file['Amount'])
        calc_closing_balance = df_write_file['Amount'].sum()

        sumval = bank_closing_balance + mm_overnight
        if(round(calc_closing_balance, 2) != sumval):
            df_exceptions = pd.concat([df_exceptions, pd.DataFrame({
                'Bank Reference ID': [id],
                'Closing_MM': [bank_closing_balance + mm_overnight],
                'Calc Closing': [calc_closing_balance]
            })], ignore_index=True)
            exceptions_bool = True
        
        timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

        df_write_file.to_excel(os.getcwd() + '\\Output\\' + str(id) + ' ' + timestamp + '.xlsx', sheet_name="Bank Transcations")
        df_mapping.loc[(df_mapping['Bank Ref ID'] == id), ['Starting_Balance']] = calc_closing_balance
    #END FOR

df_mapping.to_excel(new_mapping_path)

if not df_exceptions.empty:
    df_exceptions.to_excel(os.getcwd() + '\\Output\\EXCEPTIONS ' + timestamp + '.xlsx')