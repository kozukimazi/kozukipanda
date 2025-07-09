import pandas as pd 
import numpy as np

#read excel in the directory
df = pd.read_excel('Consolidado Modulo 1 Problema_v3.xlsx')
#print(df.columns)

# Get the position of "nombre"
nombre_pos = df.columns.get_loc("Nombre del estudiante") - 1  # +1 to place it after "nombre"
#excel
# Insert a new column (e.g., "nombre_normalizado") next to "nombre"
#.extract the last numeric part, end of string 
df.insert(nombre_pos, "Rut falso", df["Nombre del estudiante"].str.extract(r"(\d+)$"))  # Example: uppercase names

# Clean the raw names (replace underscores and remove trailing numbers)
df["nombre limpio"] = (
    df["Nombre del estudiante"]
    .str.replace("_", " ")                # Replace _ with spaces
    .str.replace(r"\s*\d+$", "", regex=True)  # Remove trailing numbers
    .str.strip()                          # Trim whitespace
)

df["rut"] = df["Persona"].str.extract(r"(\d{7,8}-[\dKk])$")  

df["solonombre"] = df["Persona"].str.replace(r"\s*\d{7,8}-[\dKk]$", "", regex=True).str.strip()
df["nombreff"] = (
    df["solonombre"]
    .str.replace(",", " ")
    .str.strip() )

# Get the position of "Nombre del etudiante" column
nombre_position = df.columns.get_loc("Nombre del estudiante") -2
nombrep = df.columns.get_loc("Persona") +1
# Insert the new column BEFORE "Nombre del etudiante"
df.insert(nombre_position, "nombre_limpio", df["nombre limpio"])
df.insert(nombrep, "Rutf", df["rut"])
df.insert(nombrep+1, "Solonombref", df["nombreff"])

# Drop the temporary "nombre limpio" column (optional)
df.drop("nombre limpio", axis=1, inplace=True)
df.drop("rut", axis=1, inplace=True)
df.drop("Entrega", axis=1, inplace=True)
df.drop("solonombre", axis=1, inplace=True)
df.drop("nombreff", axis=1, inplace=True)
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("modulo1.xlsx", engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name="Sheet1",index = False, startrow=1, header=False)

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
worksheet.set_column('M:M', 10)
worksheet.set_column('N:N', 40)
worksheet.set_column('AA:AA',60)
worksheet.set_column('Z:Z', 60)
worksheet.set_column('B:B', 25)
worksheet.set_column('C:C', 25)
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
