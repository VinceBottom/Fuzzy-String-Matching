import pandas as pd
import os
import openpyxl
import string
import difflib
os.chdir("C:\\Directory")
##add in prior year dataset
fulldf = pd.read_excel("DataSet1.xlsx", engine="openpyxl")
##add in created merged dataset
workingdf =  pd.read_excel("DataSet2.xlsx", engine="openpyxl")
##subsection
##optional drop duplicates
fulldf.drop_duplicates(subset=['User ID'], keep='last',inplace=True)
##strip whitespace from names for increased accuracy of match
fulldf['First Name'] = fulldf['First Name'].str.strip()
fulldf['Last Name'] = fulldf['Last Name'].str.strip()
##make name variable column- change name columns to be the same format (First Name space Last Name)
fulldf['Name'] = fulldf['First Name'] + ' ' + fulldf['Last Name']
##clean names
fulldf['Name'] = fulldf['Name'].str.lower()
##make strings
fulldf['Name'] = fulldf['Name'].astype("string")
##make datetimes
workingdf['DOB'] = pd.to_datetime(workingdf['DOB'])
fulldf['DOB'] = pd.to_datetime(fulldf['DOB'])
def close_match(x, df):
    df_clean = df[df.notna()]
    ##recommend match ratio of .8 or higher
    matches = difflib.get_close_matches(x, df_clean.values, n=2, cutoff=0.83)
    if matches:
        ##creates entry where there is a match according to the cutoff value
        return matches[0]
    else:
        return None
##creates column to log best match, if existing
workingdf['best_match'] = workingdf['Name'].apply(lambda x: close_match(x, fulldf['Name']))
print(workingdf['best_match'].value_counts(dropna=False))
merged_df = pd.merge(workingdf,fulldf,how='outer',left_on=['best_match','DOB'],right_on=['Name','DOB'], indicator="Merge")
merged_df2 = merged_df[merged_df['Merge']!= 'right_only']
merged_df2.to_excel("finalmerge.xlsx")
