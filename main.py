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

# Tabs to update based on the "Informations générales" filtering
tabs_to_update = [
    "Education", "Place de la femme", "Culture", "Activités et revenus",
    "Alimentation", "Agriculture", "Elevage", "Pêche", "Commerce",
    "Environnement", "Besoins", "Infrastructures", "Santé",
    "Eau et assainissement", "Energie"
]

# Load the "Informations générales" tab to get the IDs to keep
df_info_gen = pd.read_excel(original_file_path, sheet_name='Informations générales')
IDs_to_keep = df_info_gen[df_info_gen['Fokontany'].isin(villages_to_keep)]['ID_Ménage'].tolist()

# Initialize Excel writer
with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
    # For each tab, including "Informations générales", filter and save
    for sheet_name in ['Informations générales'] + tabs_to_update:
        df = pd.read_excel(original_file_path, sheet_name=sheet_name)
        # Filter rows based on the IDs_to_keep
        df_filtered = df[df['ID_Ménage'].isin(IDs_to_keep)]
        df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)

print("The modified file has been saved with updates to the specified tabs.")


#test this change
