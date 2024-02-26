import pandas as pd

# Explicit list of village names to keep
villages_to_keep = [
    "Ankotika", "Antrema", "Ampamakia", "Marosely", "Ambolipamba",
    "Andranomena", "Bobasakoa", "Ambalakida", "Ambalabe Ouest",
    "Antafiantsivakina", "Antsaonjo", "Anjiajia", "Beloy", "Ankerika Nord",
    "Ambiky", "Ambiky", "Andampy", "Ambalarano", "Komaronga", "Tsinjoarivo"
]

# The path to your original Excel file
original_file_path = '/Users/thomasdurand/Desktop/BONDY/diana_excel/donnees_DIANA_SOFIA.xlsx'

# The path for the new modified Excel file
new_file_path = '/Users/thomasdurand/Desktop/BONDY/diana_excel/donnees_DIANA_SOFIA_filtered.xlsx'

# Tabs to update based on the "Informations générales" filtering, excluding "Informations générales"
tabs_to_update = [
    "Education", "Place de la femme", "Culture", "Activités et revenus",
    "Alimentation", "Agriculture", "Elevage", "Pêche", "Commerce",
    "Environnement", "Besoins", "Infrastructures", "Santé",
    "Eau et assainissement", "Energie"
]

# Load the "Informations générales" tab to get the IDs and Fokontany to keep
df_info_gen = pd.read_excel(original_file_path, sheet_name='Informations générales')
IDs_to_keep = df_info_gen[df_info_gen['Fokontany'].isin(villages_to_keep)]['ID_Ménage'].tolist()
# Create a mapping of ID_Ménage to Fokontany for later use
id_to_fokontany = df_info_gen.set_index('ID_Ménage')['Fokontany'].to_dict()

# Initialize Excel writer
with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
  
  <<<<<<< Sub
    # Save "Informations générales" tab as is
    df_info_gen.to_excel(writer, sheet_name='Informations générales', index=False)

    # Process each specified tab
    for tab in tabs_to_update:
        if tab in pd.ExcelFile(original_file_path).sheet_names:
            df_tab = pd.read_excel(original_file_path, sheet_name=tab)
            # Filter rows based on the IDs_to_keep
            df_filtered = df_tab[df_tab['ID_Ménage'].isin(IDs_to_keep)]
            # Add Fokontany column from mapping
            df_filtered.insert(1, 'Fokontany', df_filtered['ID_Ménage'].map(id_to_fokontany))
            df_filtered.to_excel(writer, sheet_name=tab, index=False)

print("The modified file has been saved with the 'Fokontany' column added to the specified tabs.")

#This should work


