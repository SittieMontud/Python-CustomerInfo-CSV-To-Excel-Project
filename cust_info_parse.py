import pandas as pd
import numpy as np
import openpyxl as xl

#Reading CSV file
source_file = "./datasets/source_crm/cust_info.csv"
df=pd.read_csv(source_file)

#-------- cst_id --------
#Removing rows without cst_id value
df_cleaned = df.dropna(subset=['cst_id'])
#Setting cst_id as int type
df_cleaned['cst_id'] = df_cleaned['cst_id'].astype(int)


#-------- cst_firstname and cst_lastname --------
#Combine cst_firstname and cst_lastname
df_cleaned['cst_name'] = df_cleaned['cst_firstname']+' '+df_cleaned['cst_lastname']
#Fill empty rows with NA
df_cleaned['cst_name'] = df_cleaned['cst_name'].fillna('NA')
#Clean spaces and justfiy left
df_cleaned['cst_name'] = df_cleaned['cst_name'].str.strip().str.ljust(20)
#Removing columns cst_firstname and cst_lastname
df_cleaned = df_cleaned.drop(columns=['cst_firstname', 'cst_lastname'])

#-------- cst_marital_status --------
#Fill empty rows with NA
df_cleaned['cst_marital_status'] = df_cleaned['cst_marital_status'].fillna('NA')
#Change values 
df_cleaned['cst_marital_status'] = np.where(df_cleaned['cst_marital_status'] == 'M', 'Married', df_cleaned['cst_marital_status'])
df_cleaned['cst_marital_status'] = np.where(df_cleaned['cst_marital_status'] == 'S', 'Single', df_cleaned['cst_marital_status'])
#Clean spaces and justfiy left
df_cleaned['cst_marital_status'] = df_cleaned['cst_marital_status'].str.strip().str.ljust(20)

#-------- cst_gndr --------
#Fill empty rows with NA
df_cleaned['cst_gndr'] = df_cleaned['cst_gndr'].fillna('NA')
#Change values 
df_cleaned['cst_gndr'] = np.where(df_cleaned['cst_gndr'] == 'M', 'Male', df_cleaned['cst_gndr'])
df_cleaned['cst_gndr'] = np.where(df_cleaned['cst_gndr'] == 'F', 'Female', df_cleaned['cst_gndr'])
#Clean spaces and justfiy left
df_cleaned['cst_gndr'] = df_cleaned['cst_gndr'].str.strip().str.ljust(20)

#-------- Rearrange columns --------
new_order = ['cst_id','cst_key','cst_name','cst_create_date','cst_gndr','cst_marital_status']
df_cleaned = df_cleaned[new_order]

#Transfer to excel
data_to_excel = pd.ExcelWriter('CustInfo.xlsx')
df_cleaned.to_excel(data_to_excel)
data_to_excel.close()

print(df_cleaned)