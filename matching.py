# Define maximum number of applicants per option and maximum number of options per applicant
MAXIMUM_PER_OPTION = 30
MAXIMUM_PER_APPLICANT = 8
SAVE_PATH = 'matching_results'


import pandas as pd
import os
import shutil
import requests
from ID import spread_sheet_id


# Import data into pandas table
# Read the Google Sheet data into a pandas DataFrame
df = pd.read_csv(f'https://docs.google.com/spreadsheets/d/{spread_sheet_id}/export?format=csv')
df = df[df['Ich werde am FKT 2024 am 27.03.2023...'] ==
"... teilnehmen und möchte Einzelgespräche / Workshops besuchen. (Weiter geht's!)"]

# Randomize order of rows
df = df.sample(frac=1).reset_index(drop=True)

# Fill a new table with all applicants and options, default to False
identifying_columns = ['Deine E-Mail Adresse:', 'Vorname', 'Nachname', 'Aktueller Studienabschnitt',
                                   'Dein Studiengang', 'Deine Universität',
                                   'BEA-Jahrgang\nBsp.: 24 (kein Punkt, Anführungszeichen oder sonst was)',
                                   'Bitte lade hier Deinen aktuellen Lebenslauf als PDF-Datei hoch!']
matching = df[identifying_columns].copy()
options = ['Auswärtiges Amt', 'Boston Consulting Group', 'Allianz ONE', 'Bertelsmann', 'Brainlab', 'FarmInsect', 'Pelion Green Future',
           'McKinsey & Co.', 'MunichRE', 'Orcan Energy AG', 'Ritzenhöfer & company', 'SAX Power', 'SelectCode', 'Siemens Advanta Consulting',
           'TNG Consulting', 'Wacker Chemie']
non_options = ['CARIAD (noch nicht sicher)', 'Mich interessiert keine weitere Firma. Mit Auswahl dieser Option laufe ich Gefahr, weniger oder kein Gespräch zu führen.']

for option in options:
    matching['Einzelgespräch ' + option] = False

presentations = ['Allianz ONE', 'Auswärtiges Amt', 'Bertelsmann', 'Boston Consulting Group', 'FarmInsect',
                 'Pelion Green Future', 'Ritzenhöfer & company']
non_option_presentation = ['Ich habe kein Interesse an einer weiteren Präsentation und bin mir bewusst, dass ich so meine Chancen auf die Teilnahme an einer Präsentation mindere.']

for presentation in presentations:
    matching['Präsentation ' + presentation] = False

matching['Workshop The Boston Consulting Group'] = False

relevant_columns = []
for option in options:
    relevant_columns.append('Einzelgespräch ' + option)
for presentation in presentations:
    relevant_columns.append('Präsentation ' + presentation)
relevant_columns.append('Workshop The Boston Consulting Group')


# For each entry of first table
# If not one of the maxes reached: set value to true
# Randomize order of rows after each pass
for choice in range(1, 9):
    for index, row in df[['Deine E-Mail Adresse:', f'Deine {choice}. Präferenz für ein Einzelgespräch:']].iterrows():
        if row[f'Deine {choice}. Präferenz für ein Einzelgespräch:'] in non_option_presentation:
            continue
        if row[f'Deine {choice}. Präferenz für ein Einzelgespräch:'] in options and matching[f'Einzelgespräch {row[f"Deine {choice}. Präferenz für ein Einzelgespräch:"]}'].sum() < MAXIMUM_PER_OPTION and matching.loc[matching['Deine E-Mail Adresse:'] == row['Deine E-Mail Adresse:'], relevant_columns].values.ravel().sum() < MAXIMUM_PER_APPLICANT:
            matching.loc[matching['Deine E-Mail Adresse:'] == row['Deine E-Mail Adresse:'], f'Einzelgespräch {row[f"Deine {choice}. Präferenz für ein Einzelgespräch:"]}'] = True
    # there are only 7 presentations
    if choice == 8: continue
    for index, row in df[['Deine E-Mail Adresse:', f'Deine {choice}. Präferenz für eine Präsentation ']].iterrows():
        if row[f'Deine {choice}. Präferenz für eine Präsentation '] in non_options:
            continue
        if row[f'Deine {choice}. Präferenz für eine Präsentation '] in options and matching[f'Präsentation {row[f"Deine {choice}. Präferenz für eine Präsentation "]}'].sum() < MAXIMUM_PER_OPTION and matching.loc[matching['Deine E-Mail Adresse:'] == row['Deine E-Mail Adresse:'], relevant_columns].values.ravel().sum() < MAXIMUM_PER_APPLICANT:
            matching.loc[matching['Deine E-Mail Adresse:'] == row['Deine E-Mail Adresse:'], f'Präsentation {row[f"Deine {choice}. Präferenz für eine Präsentation "]}'] = True

    # BCG Workshop has to be done only once
    if choice == 1:
        for index, row in df[
            ['Deine E-Mail Adresse:', 'Interesse am Workshop von The Boston Consulting Group']].iterrows():
            if row[
                'Interesse am Workshop von The Boston Consulting Group'] == 'Ich habe an einer Workshopteilnahme Interesse.' and \
                    matching.loc[matching['Deine E-Mail Adresse:'] == row[
                        'Deine E-Mail Adresse:'], relevant_columns].values.ravel().sum() < MAXIMUM_PER_APPLICANT and \
                    matching['Workshop The Boston Consulting Group'].sum() < MAXIMUM_PER_OPTION:
                matching.loc[matching['Deine E-Mail Adresse:'] == row[
                    'Deine E-Mail Adresse:'], 'Workshop The Boston Consulting Group'] = True

    # Randomize order of rows
    df = df.sample(frac=1).reset_index(drop=True)


# Create a table and a folder for each option, save CVs and table in folder
if os.path.exists(SAVE_PATH):
    # Remove the directory
    shutil.rmtree(SAVE_PATH)
os.mkdir(SAVE_PATH)
os.chdir(SAVE_PATH)
for option in relevant_columns:
    print(f'{option}: {matching[option].sum()}')
    os.mkdir(option)
    os.chdir(option)
    option_df = matching.loc[matching[option] == True, identifying_columns]
    option_df.to_csv(f'Kandidaten {option}.csv')
    for index, row in option_df.iterrows():
        url = row["Bitte lade hier Deinen aktuellen Lebenslauf als PDF-Datei hoch!"]
        # Modify the URL to point directly to the file download
        file_id = url.split('id=')[1]
        download_url = f'https://drive.google.com/uc?export=download&id={file_id}'
        response = requests.get(download_url)

        # Ensure the request was successful
        if response.status_code == 200:
            # Open the file in write mode and write the content
            with open(f'{row["Nachname"]}{row["Vorname"]}.pdf', 'wb') as file:
                file.write(response.content)
    os.chdir('..')
    shutil.make_archive(option, 'zip', option)

for index, row in matching.iterrows():
    print(f'{row["Vorname"]} {row["Nachname"]}: {row[relevant_columns].sum()}')
