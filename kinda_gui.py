import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
import pandas as pd
import pycountry_convert as pc
from pathlib import Path


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
    

class EconomicFreedomConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor de Dados de Liberdade Econômica")
        self.root.geometry("800x600")
        
        # Instruções
        tk.Label(self.root, text="Fluxo de Trabalho:", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Passo 1
        tk.Label(self.root, text="1. Baixe os dados do Fraser Institute:", font=('Arial', 12)).pack(pady=5)
        link = tk.Label(self.root, text="Acessar Fraser Institute Dataset", fg="blue", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: webbrowser.open("https://www.fraserinstitute.org/economic-freedom/dataset?geozone=world&year=2022&page=dataset&min-year=2&max-year=0&filter=0"))
        
        # Passo 2
        tk.Label(self.root, text="2. Selecione o arquivo baixado:", font=('Arial', 12)).pack(pady=10)
        tk.Button(self.root, text="Escolher Arquivo", command=self.select_input_file).pack()
        
        # Passo 3
        tk.Label(self.root, text="3. Selecione onde salvar o arquivo convertido:", font=('Arial', 12)).pack(pady=10)
        tk.Button(self.root, text="Escolher Local de Destino", command=self.convert_file).pack()
        
        # Instruções Tableau
        tk.Label(self.root, text="\nInstruções para importar no Tableau:", font=('Arial', 12, 'bold')).pack(pady=10)
        instructions = """
        1. Abra o Tableau
        2. Clique em 'Conectar a dados'
        3. Selecione 'Microsoft Excel'
        4. Navegue até o arquivo 'FonteConvertida.xlsx'
        5. Importe os dados
        """
        tk.Label(self.root, text=instructions, justify=tk.LEFT).pack(pady=10)
        
        self.input_file_path = None
        
    def select_input_file(self):
        self.input_file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if self.input_file_path:
            messagebox.showinfo("Sucesso", "Arquivo selecionado com sucesso!")
    
    def convert_file(self):
        if not self.input_file_path:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo primeiro!")
            return
            
        output_path = filedialog.asksaveasfilename(
            title="Salvar arquivo convertido",
            defaultextension=".xlsx",
            initialfile="FonteConvertida.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if output_path:
            try:
                # Use the existing code but with dynamic input/output files
                df = pd.read_excel(self.input_file_path, skiprows=4)
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

                df.to_excel(output_path, index=False)
                messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao converter arquivo: {str(e)}")

if __name__ == "__main__":
    app = EconomicFreedomConverter()
    app.root.mainloop()