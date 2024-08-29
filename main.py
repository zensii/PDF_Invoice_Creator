import pandas as pd
import glob

filepaths = glob.glob('xls_files/*.xlsx')

for file in filepaths:
    df = pd.read_excel(file)
    print(df)