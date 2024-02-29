# import packages
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load the excel file (all the sheets)
excel_sheets_df = pd.read_excel('C:/Users/subhr/Documents/Python Scripts/EDA/Superstore_sales.xlsx',
                    sheet_name=['Orders','Returns','Users'])

# print (excel_sheets_df['Orders'])
orders_df = excel_sheets_df['Orders']
returns_df = excel_sheets_df['Returns']
users_df = excel_sheets_df['Users']

# print(orders_df)
# print(returns_df)
# print(users_df)

# Convert the individual sheet data into Pandas DataFrames
Orders = pd.DataFrame(orders_df)
Returns = pd.DataFrame(returns_df)
Users = pd.DataFrame(users_df)

# Check if there are any duplicate records in the DataFrames
orders_dup = Orders.duplicated().sum()
returns_dup = Returns.duplicated().sum()
users_dup = Users.duplicated().sum()

# Print the no. of duplicate records in each of the DataFrames
print('There are %d duplicate records in Orders' % (orders_dup))
print('There are {} duplicate records in Returns'.format(returns_dup))
print('There are {} duplicate records in Users'.format(users_dup))

# Check for the missing values in the DataFrames
print(Orders.shape)

#orders_missing = Orders.isnull().sum()
#print(orders_missing)

# identify the data types for each of the attributes
# Check for the outliers
# Check for Co-relations between various numerical attributes
