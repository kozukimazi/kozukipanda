import pandas as pd 
import numpy as np
#from fuzzywuzzy import fuzz

#read excel in the directory
df = pd.read_excel('notas.xlsx')

#function of conversion for the modules
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

#aqui me defino la función que pondrá las notas en el excel
def check_phrases(row):
    """
    Checks the phrases in a string separated by '|' and returns:
    - n=0 if phrase1 is "hola" and phrase2 is "chao"
    - n=1 if phrase1 is "hello" and phrase2 is "goodbye"
    - n=2 for any other combination
    """
    input_string = row['Observaciones']
    info1 = row['Info1']
    info2 = row['Info2']
    info3 = row['Info3']
    # Split the string by the delimiter '|'
    phrases = [phrase.strip() for phrase in input_string.split('|')]
    if len(phrases) == 1:
        print("No delimiter found, using full string as phrase1")
        phrase1 = phrases[0]
        phrase2 = None 
    else:
        phrase1, phrase2 = phrases

    phrase11 = "Nivel Básico en Números y Proporcionalidad: Debes realizar el Módulo 2"
    phrase12 = "Nivel Medio en Números y Proporcionalidad: Te recomendamos realizar el Módulo 2"
    phrase13 = "Nivel Avanzado en Números y Proporcionalidad: Te eximes del Módulo 2"

    phrase21 = "Nivel Básico en Álgebra y Funciones: Debes realizar el Módulo 3"
    phrase22 = "Nivel Medio en Álgebra y Funciones: Te recomendamos realizar el Módulo 3"
    phrase33 = "Nivel Avanzado en Álgebra y Funciones: Te eximes del Módulo 3"

    Info10 = "No rindio"
    Info11 = "Dada info1" 

    Info20 = "No rindio"
    Info21 = "Dada info2"

    Info30 = "No rindio"
    Info31 = "Dada info3"

    #esto tendré que verlo bien, pero es la función que añade la columna final
    if phrase1 == phrase11 and phrase2 == phrase21 and info1 == Info11 and info2 == Info21 and info3 == Info31:
        return 'Nota 1| Nota 2| Nota 3'
    elif phrase1 == phrase11 and phrase2 == phrase21 and info1 == Info11 and ( (info2 == Info21 and info3 == Info30) or (info2 == Info20 and info3 == Info31) ):
        return 'Nota1| Nota 2'
    else:
        return 'Value3'


#function to know if "jose peres" is in "jose peres duarte"
def is_name_contained(short_name, long_name):
    return short_name in long_name  

#read frame 2 and check a list 
df2 = pd.read_excel('lista_entregas.xlsx')
names_list = df2['Persona'].tolist()
#read frame 3 and check list
df3 = pd.read_excel('info2.xlsx')
names_list2 = df3['Persona'].tolist()
#read frame4 and check list
df4 = pd.read_excel('info3.xlsx')
names_list3 = df4['Persona'].tolist()

def names(name,x=names_list):

    for xs in x:
        if (is_name_contained(name,xs)):
            return 'Dada info1'
              
    return 'No rindio'

def names2(name,x=names_list2):

    for xs in x:
        if (is_name_contained(name,xs)):
            return 'Dada info2'
              
    return 'No rindio'

def names3(name,x=names_list3):

    for xs in x:
        if (is_name_contained(name,xs)):
            return 'Dada info3'
              
    return 'No rindio'

#fuzzy es muy preciso, tanto que no lo necesito
#def check_fuzzy_match(name, threshold=80):
    #for candidate in names_list:
      #  if fuzz.ratio(name.lower(), candidate.lower()) >= threshold:
     #       return True
    #return False

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

#aplico la conversion por elemento de cada columna        
#if 'P1 (50%)' in df.columns:
 #   df['Modulo 2'] = df['P1 (50%)'].apply(conversion2) 

#if 'P2 (50%)' in df.columns:
 #   df['Modulo 3'] = df['P2 (50%)'].apply(conversion2) 

names_list = df2['Persona'].tolist()

df['Info1'] = df['Persona'].apply(names)
df['Info2'] = df['Persona'].apply(names2)
df['Info3'] = df['Persona'].apply(names3)
#df['Info1'] = df['Persona'].apply(check_fuzzy_match)

print(df)
print(df.columns)
#df = df.drop(columns=['Observaciones'], errors='ignore')
print(df.columns)
print(df.shape)  

print(df['Info1'].head())



#print(df.columns)
#print(df['Persona'].head())  
#cuenta el numero de veces de un elemento especifico  
conteo1 = df['Info1'].value_counts().get('Dada info1', 0)
conteo2 = df['Info2'].value_counts().get('Dada info2', 0)
conteo3 = df['Info3'].value_counts().get('Dada info3', 0)
print(conteo1)
print(conteo2)
print(conteo3)
#print(names_list[1])

#here we are going to use format
#df.to_excel('notasf1.xlsx', index=False)  # index=False avoids saving row numbers

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("notasf1.xlsx", engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name="Sheet1",index = False, startrow=1, header=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# Set the column width and format.
worksheet.set_column('A:A', 5)
worksheet.set_column('B:B', 40)
worksheet.set_column('C:C', 15)
worksheet.set_column('D:D', 15)
worksheet.set_column('E:E', 8)
worksheet.set_column('F:F', 8)
worksheet.set_column('G:G', 8)
worksheet.set_column('H:H', 60)
worksheet.set_column('I:I', 10)
worksheet.set_column('J:J', 10)
worksheet.set_column('K:K', 10)
worksheet.set_column('L:L', 10)

 # Create a border format
border_format = workbook.add_format({
        'border': 1,  # 1 = thin border
        'border_color': 'black'
})
    
# Apply borders to all cells with data
max_row, max_col = df.shape
worksheet.conditional_format(0, 0, max_row, max_col, 
                               {'type': 'no_blanks',
                                'format': border_format})

#add a header format
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': "#8A907E",
    'border': 1})

# Write the column headers with the defined format.
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num , value, header_format)

# Close the Pandas Excel writer and output the Excel file.
writer.close()


