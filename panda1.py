import pandas as pd 
import numpy as np

#read excel in the directory
df = pd.read_excel('notas.xlsx')
#print(df)
#print(df.shape)
#print(df['P1 (50%)'].head())
#print(df.columns)

def conversion(x):
    if (type(x)==float):
            #el error esta aquí, porque le entrego toda la columna y la vuelve "No rinde"
        return 'No rinde'
    else:  
        #print(float(x.replace('%', '')))  
        if (0< float(x.replace('%', '')) <= 40 ):
            return  'Debe'
        elif (40 <float(x.replace('%', '')) <= 80):
            return 'Recomienda'
        else:
            return 'Exime'

if 'P1 (50%)' in df.columns:
    df['Modulo 2'] = df['P1 (50%)'].apply(conversion)    
print(df)
print(df.columns)
df = df.drop(columns=['Observaciones'], errors='ignore')
print(df.columns)
print(df.shape)      

df.to_excel('notas1.xlsx', index=False)  # index=False avoids saving row numbers