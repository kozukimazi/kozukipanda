import pandas as pd 
import numpy as np


def notaschile(p):
    nmax = 7.00
    nmin = 1.00
    napr = 4.00
    porc = 0.6
    pmax = 56
    if (p<porc*pmax):
        a = (napr-nmin)*(p/(porc*pmax)) + nmin
        formatted = "{:.2f}".format(a)
        return a 
    else:
        a = (nmax-napr)*((p-porc*  pmax)/((1-porc)*pmax)) + napr
        formatted = "{:.2f}".format(a)
        return a 
    
#read excel in the directory
dffono = pd.read_excel('fono.xlsx')
dftera = pd.read_excel('terapia.xlsx')
dfnutri = pd.read_excel('nutri.xlsx')

dfdes =  pd.read_excel('DesarrolloP3.xlsx', sheet_name="Desarrollo")



dffono = dffono.drop("Unnamed: 3",axis = 1)
dftera = dftera.drop("Unnamed: 3",axis=1)
dfnutri = dfnutri.drop("Unnamed: 3",axis = 1)



#dffono = dffono.loc[:,~dffono.columns.duplicated()]
#dfnutri = dfnutri.loc[:,~dfnutri.columns.duplicated()]
#dftera = dftera.loc[:,~dftera.columns.duplicated()]

# merge the two columns into one
#dffono['nombre'] = pd.concat([dffono['Nombre'], dffono['N']])
#dftera['nombre'] = pd.concat([dftera['Nombre'], dftera['N']])
#dfnutri['nombre'] = pd.concat([dfnutri['Nombre'], dfnutri['N']])


dffono['Carrera'] = "FO"
dftera['Carrera'] = "TO"
dfnutri['Carrera'] = "NUT"

dftot = pd.concat([dffono, dftera, dfnutri], ignore_index=True)
dftot = dftot.drop('N°',axis = 1)
dfdes = dfdes.drop('Corrector ',axis = 1)

# Convert "Nombre" in A to a list (this defines the target order)
orden = dftot["Nombre"].tolist()

# Step 1: Set "Persona" as the index of B (for alignment)
df_b_sorted = dfdes.set_index("Persona")

# Step 2: Reindex B using the order from A's "Nombre"
df_b_sorted = df_b_sorted.reindex(orden)

# Step 3: Reset index to bring "Persona" back as a column
df_b_sorted = df_b_sorted.reset_index()

# Get the position of "Nombre" (+1 to insert behind it)
nombre_position = dftot.columns.get_loc("Nombre") + 1
nombre0 = dftot.columns.get_loc("Buenas")+1
# Insert the column (e.g., "Edad_from_B") into DataFrame A
#dftot.insert(nombre_position, "personaB", df_b_sorted["Persona"])
dftot.insert(nombre_position, "RUT", df_b_sorted["RUT"])
dftot.insert(nombre0, "P1A", df_b_sorted["P1A"])
dftot.insert(nombre0+1, "P1B", df_b_sorted["P1B"])
dftot.insert(nombre0+2, "P1C", df_b_sorted["P1C"])
dftot.insert(nombre0+3, "P2A", df_b_sorted["P2A"])
dftot.insert(nombre0+4, "P2B", df_b_sorted["P2B"])
dftot.insert(nombre0+5, "P2C", df_b_sorted["P2C"])
dftot.insert(nombre0+6, "Total", df_b_sorted["Total"])
#dftot.insert(nombre0+7, "Forma", df_b_sorted["Forma"])

dftot['Puntaje'] = dftot['Total'] + 2*dftot['Buenas']
dftot['Nota 60 %'] = dftot['Puntaje'].apply(notaschile)


# Check if the order matches A
print(df_b_sorted["Persona"].equals(dftot["Nombre"]))  # Should return True



# Save the reordered DataFrame B to a new Excel file
df_b_sorted.to_excel("P3reorder.xlsx", index=False)

#######We create three auxiliar excels##
dffono.to_excel("fono0.xlsx",sheet_name='Sheet_name_1') 
dftera.to_excel("tera0.xlsx",sheet_name='Sheet_name_1') 
dfnutri.to_excel("nutri0.xlsx",sheet_name='Sheet_name_1') 

dftot.to_excel("fonuto0f.xlsx",sheet_name='Sheet_name_1') 

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("fonutofinal.xlsx", engine="xlsxwriter")

dftot.to_excel(writer, sheet_name="Sheet1",index = False, startrow=1, header=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# Set the column width and format.
#wwet_column('A:A', 5)
#worksheet.set_column('B:B', 40)
##worksheet.set_column('C:C', 15)
#worksheet.set_column('D:D', 15)
#worksheet.set_column('E:E', 8)
#worksheet.set_column('F:F', 8)
#worksheet.set_column('G:G', 8)
worksheet.set_column('H:H', 30)
worksheet.set_column('I:I', 30)
worksheet.set_column('J:J', 30)
worksheet.set_column('K:K', 30)
worksheet.set_column('L:L', 30)
 # Create a border format
border_format = workbook.add_format({
        'border': 1,  # 1 = thin border
        'border_color': 'black'
})
    
# Apply borders to all cells with data
max_row, max_col = dftot.shape
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
for col_num, value in enumerate(dftot.columns.values):
    worksheet.write(0, col_num , value, header_format)

# Close the Pandas Excel writer and output the Excel file.
writer.close()



