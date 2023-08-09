import pandas as pd

# Load the path to the missing numbers sheet
df_missing = pd.read_excel(r"C:\Users\lhollestelle\Downloads\Staff Phone Number Missing.xlsx")

# Load the path to the all phone numbers sheet
df_all = pd.read_excel(r"C:\Users\lhollestelle\Downloads\All Phone Numbers.xlsx")


# loop through both files to find matches and update the work_phone
missing_index = -1
for missing_name in df_missing.loc[:,'NAME']:
    missing_index += 1
    real_index = -1
    for real_name in df_all.loc[:,'NAME']:
        real_index += 1
        if str(missing_name) == str(real_name).upper():
            phoneNumber = df_all.loc[real_index, 'EXT']
            df_missing.loc[missing_index, 'WORK_PHONE'] = int(phoneNumber)
            break
print(df_missing)

# Save the updated destination sheet, uncomment the next line when you want to edit the real excel file
df_missing.to_excel(r"C:\Users\lhollestelle\Downloads\Staff Phone Number Missing.xlsx", index=False)
