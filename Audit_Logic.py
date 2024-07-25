import pandas as pd
import numpy as np
import datetime as dt
from IPython import embed

def excel_read():
    df=pd.read_excel(r"/home/user/Downloads/33. AWS_DFS_RPF_List28_24072024_Cycle_33.xlsx",sheet_name='Sheet1')
    return df



def header_map(df):
    column_headers = ['Campaign Name', 'Campaign ID', 'Member ID', 'Email', 'First Name', 'Last Name', 'Job Title', 'Job Level', 'Job Function',
                      'Company', 'Address', 'City', 'State', 'Zip', 'Country', 'Phone', 'AlternateNo', 'Ext', 'Time Zone', 'Employee Size',
                      'Company Revenue', 'Industry', 'Priority', 'Owner', 'Original Owner', 'Website', 'Remarks', 'File Name', 'Client / Net New', 'AC ID',    
                      'Contact ID', 'Phone Remarks', 'Mobile', 'Direct', 'Ext1', 'Location Phone', 'HQ Phone 1', 'HQ Phone 2', 'HQ Phone 3',
                      'HQ Phone 4', 'HQ Phone 5', 'Toll Free', 'Lusha1', 'Lusha Country 1', 'Lusha2', 'Lusha Country 2', 'Person Convert Link Li',    
                      'Person Link Li', 'Person Other Link', 'Company Link', 'Mobile Phone Source', 'Direct Phone Source', 'Location Phone Source',
                      'Phone Source1', 'Phone Source2', 'Phone Source3', 'Phone Source4', 'Phone Source5', 'Address Source',    
                      'Employees Source', 'Revenue Source', 'Industry Source', 'Lusha  Email 1', 'Lusha  Email 2', 'Lusha Phone Remarks',    
                      'Lusha RA Name', 'Li contact Location', 'ZB Result', 'ZB Sub status', 'ZB SMTP Provider', 'Strikeiron reasonDescription',
                      'Strikeiron Result', 'User validation Email Remark', 'User validation  Valid/Invalid', 'SendGrid Result', 'SendGrid Type Result',
                      'Exchange Result', 'Email Status', 'Contact per Company', 'Country1', 'File', 'Remarks1', 'Category', 'Category 2', 'Match Catrgory',
                      'Match State', 'Type', 'Description', 'rating', 'reviewCount', 'Source Link', 'Source', 'Remarks2', 'Dialled Number',
                      'New Contactname if found', 'Contact Verified(Y/N)', 'Company Verified (Y/N)', 'Calling Remarks', 'New phonenumber',
                      'Calling RA Name', 'Date', 'File Name1', 'File Date', 'DB_File_Name', 'For Campaign', 'Person ID', 'Direct Phone Number',
                      'Title Keyword', 'Function', 'Priority1', 'Work File', 'Website1', 'Countif', 'Other Domain', 'Country List', 'List', 'Li contact Country',
                      'Calling File Name', 'Final', 'Company LI Link', 'Estimated Revenue (Y/N)', 'company speciality', 'company industry',
                      'String', 'Calling Update Email', 'Phone Remarks1', 'Old Contact ID', 'Valid / Invalid', 'Count Per Company', 'Job Function1',
                      '(Lusha) Work email', '(Lusha) Direct email', 'Data Calling Remarks', 'US L0 Category', 'US L1 Category', 'L1 Category Website/Amazon',
                      'Store URL - Amazon', 'Contact LI Source', 'Contact Website Source', 'Contact Other Source', 'Phone Website Source', 'Phone Other Source',
                      'Generic Email', 'Email Source', 'Email Remarks', 'TAL Account ID SFDC', 'Priority2', 'TAG', 'strikeiron Phone', 'Phone State',
                      'Store Name', 'Store Address', 'Store City', 'Store State', 'Store Zip', 'Store Miles', 'Sales Subvertical', 'Sales L2 Subvertical',
                      'TAL List', 'Suppressed', 'Extra', 'Status (Large/Core)', 'PO', 'Cycle', 'Location']


    column_headers_set = set(column_headers)

    # Get the common columns between the new DataFrame (df) and the specified column headers
    common_columns = list(column_headers_set.intersection(df.columns))

    # Create a new DataFrame with only the common columns
    new_df_common = df[common_columns]

    # Get the uncommon columns in the new DataFrame
    uncommon_columns = list(column_headers_set.difference(df.columns))

    # Create a DataFrame with only the uncommon columns from the new DataFrame
    new_df_uncommon = pd.DataFrame(columns=uncommon_columns)

    # Concatenate the common and uncommon DataFrames along the columns axis
    final_df = pd.concat([new_df_common, new_df_uncommon], axis=1)
    final_df = final_df[column_headers]
    print(final_df)
    return final_df

def data_manipulate(final_df):
    Data=final_df
    Data['Email'].fillna('ia@ia.com',inplace=True)
    Data['Zip'].fillna(1234,inplace=True)
    Data.loc[Data['Job Level'].str.startswith(('D', 'd')), 'Job Level'] = 'Director Level'
    Data.loc[Data['Job Level'].str.startswith(('M', 'm')), 'Job Level'] = 'Manager Level'
    Data.loc[Data['Job Level'].str.startswith(('C', 'c')), 'Job Level'] = 'C-Level'
    Data.loc[Data['Job Level'].str.startswith(('V', 'v')), 'Job Level'] = 'VP Level'
    Data.loc[Data['Job Level'].str.startswith(('I', 'i')), 'Job Level'] = 'Individual Contributor'
    lst = ['Member ID','Email','First Name', 'Last Name', 'Job Title', 'Job Level', 'Job Function', 'Company', 'Address', 'City', 'State', 'Zip','Country',	'Phone', 'Time Zone',	'Employee Size',	'Company Revenue',	'Industry',	'Priority',	'Owner','Original Owner','Website']
    Data[lst] = Data[lst].fillna('NA')
    Data[lst] = Data[lst].replace('-', 'NA')
    special_characters = set(['!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '+', '=', '{', '}', '[', ']', ':', ';', '"', "'", '<', '>', ',', '.', '/', '\\', '|', '?', '~', '`', '°', '@'])
    for i, row in Data.iterrows():
        if all(char in special_characters for char in row['Member ID']):
            Data.at[i, 'Member ID'] = 'NA'


    print(Data)
    return Data





def main():
    excel=excel_read() 
    map_col=header_map(excel)
    mp_data=data_manipulate(map_col)
    
    output_path = '/home/user/Documents/Audit_Rules/TalkDesk/Talkdesk RPF_07242024_ZI_Cycle_v2_24_july.xlsx'
    mp_data.to_excel(output_path, index=False)
    print('Done')
    

    

if __name__ == "__main__":
    main()

