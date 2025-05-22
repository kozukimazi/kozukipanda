import pandas as pd 
import numpy as np

#read excel in the directory
df = pd.read_excel('notas.xlsx')
#print(df)
#print(df.shape)
#print(df['P1 (50%)'].head())
#print(df.columns)

def conversion2(x):
    if (type(x)==float):
            #el error esta aquí, porque le entrego toda la columna y la vuelve "No rinde"
        return 'No rindio'
    else:  
        #print(float(x.replace('%', '')))  
        if (0< float(x.replace('%', '')) <= 40 ):
            return  'Debe'
        elif (40 <float(x.replace('%', '')) <= 80):
            return 'Recomienda'
        else:
            return 'Exime'

def is_name_contained(short_name, long_name):
    return short_name in long_name  

from fuzzywuzzy import fuzz


df2 = pd.read_excel('lista_entregas.xlsx')
names_list = df2['Persona'].tolist()
def names(name,x=names_list):

    for xs in x:
        if (is_name_contained(name,xs)):
            return 'dada info1'

              
        else:
            return 'No rindio'


def check_fuzzy_match(name, threshold=80):
    for candidate in names_list:
        if fuzz.ratio(name.lower(), candidate.lower()) >= threshold:
            return True
    return False

def is_name_contained(short_name, long_name):
    return short_name in long_name

def isname(short_name, long_names=names_list):
    for long_name in long_names:
        #print(long_name)
        if(short_name in long_name):
            return True
            break
        else:
            return False
        
if 'P1 (50%)' in df.columns:
    df['Modulo 2'] = df['P1 (50%)'].apply(conversion2) 

if 'P2 (50%)' in df.columns:
    df['Modulo 3'] = df['P2 (50%)'].apply(conversion2) 

names_list = df2['Persona'].tolist()

#df['Info1'] = df['Persona'].apply(isname)
df['Info1'] = df['Persona'].apply(check_fuzzy_match)

print(df)
print(df.columns)
df = df.drop(columns=['Observaciones'], errors='ignore')
print(df.columns)
print(df.shape)  

print(df2['Persona'].head())


tot = is_name_contained("jose, perez", "jose, perez goku sayayin")
print(tot)

print(df.columns)
    
#print(names_list[1])


df.to_excel('notasf1.xlsx', index=False)  # index=False avoids saving row numbers