import pandas as pd 
import numpy as np

df = pd.read_csv('lindbladgamU',  delim_whitespace=True, names=['A1', 'A2', 'A3', 'A4','A5', 'A6', 'A7', 'A8','A9'])
print(df.head())

df.to_excel('datos.xlsx', index=False)  # index=False avoids saving row numbers