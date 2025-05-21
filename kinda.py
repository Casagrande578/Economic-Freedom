import pandas as pd
import pycountry_convert as pc


def get_region(contry_code):
    try:
        region_code = pc.country_alpha2_to_continent_code(country_2_code=contry_code)
        return pc.convert_continent_code_to_continent_name(region_code)
    except Exception as e:
        return 'N/A'
    
def get_quartile(quartile_number):
    if(quartile_number):
        if quartile_number == 1:
            return "Most Free"
        if quartile_number == 2:
            return "2nd Quartile"
        if quartile_number == 3:
            return "3rd Quartile"
        return "Least Free"
    

# Read the input Excel file, skipping the first 4 rows
input_file = 'efotw-2024-master-index-data-for-researchers-iso.xlsx'
df = pd.read_excel(input_file, skiprows=4)

# Strip leading/trailing whitespace from column names
df.columns = df.columns.str.strip()


# Create three separate dataframes for each metric
research_data = pd.concat([
    pd.DataFrame({
        'Year': df['Year'],
        'ISO Code 3': df['ISO Code 3'],
        'Countries': df['Countries'],
        'Research': 'Quartile',
        'value': df['Quartile']
    }),
    pd.DataFrame({
        'Year': df['Year'],
        'ISO Code 3': df['ISO Code 3'],
        'Countries': df['Countries'],
        'Research': 'Rank',
        'value': df['Rank']
    }),
    pd.DataFrame({
        'Year': df['Year'],
        'ISO Code 3': df['ISO Code 3'],
        'Countries': df['Countries'],
        'Research': 'Economic Freedom Summary Index',
        'value': df['Economic Freedom Summary Index']
    })
], ignore_index=True)

df = df.merge(
    research_data,
    on=['Year', 'ISO Code 3', 'Countries'],
    how='left'
)

print(df.columns)


# Rename columns using the mapping
columns_mapping = {
    'Year': 'Ano/Year',
    'World Bank Region': 'Subregião / Subregion',
    # 'World Bank Current Income Classification, 1990-Present': 'Subregião / Subregion',
    'Countries': 'País / Country',
    'Economic Freedom Summary Index': 'Indice / Index - Discrete',
    'Rank': 'Rank - World',
    'Quartile': 'Quartil / Quartile',
    'Area 1 Rank': 'Área / Area'
}

df = df.rename(columns=columns_mapping)

# Create additional columns by copying or assigning data
df['Quartile'] = df['Quartil / Quartile']
df['Quartiles - Eco Free'] = df['Quartil / Quartile']
df['Rank'] = df['Rank - World']
df['Language1'] = "English"
df['State'] = 'National'
df['Area'] = 'Índice de Liberdade Econômica / Economic Freedom Summary Index'
df['Research Code'] = ''
# df['Research'] = ''
df['indexValue - Continuous'] = df['Indice / Index - Discrete']
df['indexValue - Continuous -F'] = df['Indice / Index - Discrete']
df['Região / Region'] = df['ISO Code 2'].apply(get_region) 
df['Quartil / Quartile'] = df['Quartile'].apply(get_quartile)

# Reorder columns to match desired output
desired_columns = [
    'Language1', 'Ano/Year', 'Região / Region', 'Subregião / Subregion', 
    'País / Country', 'State', 'Area', 'Research Code', 'Research', 
    'Indice / Index - Discrete', 'Quartiles - Eco Free', 'Rank - World', 
    'Quartile', 'Rank', 'Quartil / Quartile', 'Área / Area', 
    'indexValue - Continuous -F', 'indexValue - Continuous'
]

df = df.reindex(columns=desired_columns)


# Write the DataFrame to a new Excel file
output_file = 'converted.xlsx'
df.to_excel(output_file, index=False)