import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

# Create the test data directory if it doesn't exist
os.makedirs('test_files', exist_ok=True)

# Create January data for 2024
start_date = datetime(2024, 1, 1)
dates = [start_date + timedelta(days=i) for i in range(31)]
df_jan = pd.DataFrame({
    'Timestamp': dates, 
    'Raw': np.random.randint(100, 200, 31)
})
df_jan.to_excel('test_files/2024_01_JanuaryData.xlsx', index=False)
print('Created January 2024 data file')

# Create February data
start_date = datetime(2024, 2, 1)
dates = [start_date + timedelta(days=i) for i in range(28)]
df_feb = pd.DataFrame({
    'Timestamp': dates, 
    'Raw': np.random.randint(120, 220, 28)
})
df_feb.to_excel('test_files/2024_02_FebruaryData.xlsx', index=False)
print('Created February 2024 data file')

# Create March data
start_date = datetime(2024, 3, 1)
dates = [start_date + timedelta(days=i) for i in range(31)]
df_mar = pd.DataFrame({
    'Timestamp': dates, 
    'Raw': np.random.randint(150, 250, 31)
})
df_mar.to_excel('test_files/2024_03_MarchData.xlsx', index=False)
print('Created March 2024 data file')

print('All test data created successfully')
