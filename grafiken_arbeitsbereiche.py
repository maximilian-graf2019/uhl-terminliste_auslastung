from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
import glob
import os
import seaborn as sns
import matplotlib.pyplot as plt

excel_files = glob.glob("*.xlsx")
file = excel_files[0]
if len(excel_files) > 1:
    raise ValueError('Es liegen mehrere Dateien im Ordner. Bitte lösche nicht mehr benötigte Dateien.')

print(f'Nutze die Datei "{file}"')

# get files and read first excel into df
excel_files = glob.glob("*.xlsx")
file = excel_files[0]
df = pd.read_excel(file)
print('Excel erfolgreich eingelesen.')

df.columns = df.columns.str.replace(' ', '')

# removing old files from subfolder grafiken 
try:
    filenames = os.listdir('grafiken')
except FileNotFoundError:
    os.mkdir('grafiken')
    filenames = os.listdir('grafiken')
    
for filename in filenames:
    os.unlink(f'grafiken/{filename}')

# remove whitespaces from column names
df.rename(columns=lambda x: x.strip(), inplace=True)

# get relevant columns and filter dataframe
cols = ['Auftrag'] + [col for col in df.columns if col.__contains__('KW')]
df_relevant = df[cols]
print(df_relevant.shape)
for col in df_relevant.columns:
    new_col = col.replace('\n', '_')
    df_relevant.rename(columns={col: new_col}, inplace=True)

# creating variables
arbeitsbereiche = ['Kapazität PR-Fertigung', 'Kapazität Fensterfertigung',
                   'Kapazität Türfertigung', 'Kapazität Blechfertigung', 'Kapazität Abt. Schweißen', 'Kapazität Rollen']
caps = [100, 400, 180, 50, 80, 40]
today = datetime.today()
kw = today.isocalendar().week

# printing the used Variables
print(f'Es werden die Grafiken für die KW {kw} erstellt.')
print('Folgende Arbeitsbereiche mit maximalen Kapazitäten werden verwendet:')
for i, ab in enumerate(arbeitsbereiche):
    print(i, '\t', caps[i], '\t', ab[10:])

# create df per 'Arbeitsbereich'
row_nums = list()
for ab in arbeitsbereiche:
    rownum = df_relevant.loc[df_relevant['Auftrag'] == ab].index
    row_nums.append(rownum[0])

print('Gefunden Startzeilen', row_nums, '\n')
print(df_relevant.head())

dfs = []

for i, arbeitsbereich in enumerate(arbeitsbereiche):
    df_temp = df_relevant.iloc[row_nums[i]-1:row_nums[i] +
                         1].dropna(axis=1, how='all').transpose().stack().reset_index()[1:]
    df_temp.rename(columns={'level_0': 'KW',
                'level_1': 'Legende', 0: 'value'}, inplace=True)
    # print(df_temp.info())
    row_Auslastung = row_nums[i]-1
    row_Kapazitaet = row_nums[i]
    df_temp['Legende'].replace(row_Auslastung, 'Auslastung', inplace=True)
    df_temp['Legende'].replace(row_Kapazitaet, 'Kapazität', inplace=True)
    dfs.append(df_temp) 

print('Dateien wie benötigt erstellt.', '\n')

# define function to get KW-Keys
def get_kw_names(number_of_keys: int, kw=today.isocalendar().week, year=today.year):
    # this statement gets evaluated in every normal calendar week
    if 52 - kw - number_of_keys > 0:
        if kw >= 10:
            return [str(year) + '_KW' + str(i)
                    for i in range(kw, kw + number_of_keys)]
        else:
            kws = [str(year) + '_KW0' + str(i) for i in range(kw, 10)] + [str(year) + '_KW' + str(i) for i in range(10, 10 + number_of_keys - 10 + kw)]
            return kws
    # gets evaluated, when part of the weeks go into the next year, but no more than 10
    elif number_of_keys - 51 + kw < 10:
        kws = [str(year) + '_KW' + str(i) for i in range(kw, 53)] + \
            [str(year + 1) + '_KW0' + str(i)
             for i in range(1, number_of_keys - 51 + kw)]
        return kws
    # get evaluated for the last weeks of the year, when weeks in next year is greater than 10
    else:
        kws = [str(year) + '_KW' + str(i) for i in range(kw, 53)] + \
            [str(year + 1) + '_KW0' + str(i) for i in range(1, 10)] + \
            [str(year + 1) + '_KW' + str(i)
             for i in range(10, number_of_keys - 52 + kw)]
        return kws

# define function to plot and save
def plot_abteilung(abteilung, data, capacity, kw=today.isocalendar().week):
    sns.set(rc={'figure.figsize': (30, 9)}, font_scale = 1.2)
    sns.catplot(data=data,
                kind='bar',
                x='KW',
                y='value',
                hue='Legende',
                height=6,
                aspect=2.5,
                hue_order=['Kapazität', 'Auslastung'],
                palette=sns.color_palette(['green', 'red']))
    # plt.title(
    #      f'Auslastung für {abteilung[10:]} in KW{kw}', size=16)
    plt.ylabel('Stunden')
    plt.xticks(rotation=45)
    plt.xlabel('Kalenderwochen')
    plt.axhline(capacity, c='gray')
    if abteilung[10:] == 'Abt. Schweißen':
        plt.savefig(
            f'./grafiken/Grafik_Schweissen_KW{kw}.png', 
            bbox_inches="tight")
    else:
        plt.savefig(
            f'./grafiken/Grafik_{abteilung[10:]}_KW{kw}.png', 
            bbox_inches="tight")

# get total auslastung
def get_total_auslastung(dfs):
    totals = list()
    for df in dfs:
        df_capacity = df.loc[(df.Legende == "Auslastung") & (df.KW.isin(get_kw_names(26)))].reset_index(drop=True)
        print(df_capacity.shape)
        print(df_capacity.head(30))
        sum_capacity = df_capacity['value'].sum()
        totals.append(round(sum_capacity, 2))
    return {abteilung : totals[i] for i, abteilung in enumerate(arbeitsbereiche)}

totals = get_total_auslastung(dfs)
print(totals)

# plotting and saving using the functions
for df_nr, abt in enumerate(arbeitsbereiche):
    df_current = dfs[df_nr]
    plot_abteilung(
        abteilung=abt,
        data=df_current.loc[df_current['KW'].isin(get_kw_names(26))],
        capacity=caps[df_nr])

print('PDF wird erzeugt.')

title = 'Übersicht Fertigung Auslastung + Kapazität'

def create_pdf_with_images(abteilung_list, output_filename):
    doc = SimpleDocTemplate(output_filename, pagesize=A4)

    # Prepare the list of elements for the PDF
    elements = []
    styles = getSampleStyleSheet()

    # Add title to the PDF
    title = "Auslastung der Fertigung"
    title_text = Paragraph(title, styles['Title'])
    elements.append(title_text)
    elements.append(Spacer(1, 20))

    for i, abteilung in enumerate(abteilung_list):
        # Create a heading for each techniker
        heading_text = f"Kapazität: {abteilung} (Ist:{totals[abteilung]} / Soll:{caps[i]*26} Std.)"
        heading = Paragraph(heading_text, styles['Heading2'])
        elements.append(heading)

        # Find and add the corresponding plot image to the PDF
        if abteilung[10:] == 'Abt. Schweißen':
            image_filename = f'./grafiken/Grafik_Schweissen_KW{kw}.png'
        else:
            image_filename = f'./grafiken/Grafik_{abteilung[10:]}_KW{kw}.png'

        images = glob.glob(image_filename)
        if images:
            img = Image(images[0], height=200, width=500)
            elements.append(img)

        elements.append(Spacer(1, 20))
        # Add a page break after every second techniker
        if i % 2 == 1:
            elements.append(PageBreak())

    # Build the PDF
    doc.build(elements)

# Usage example:
output_filename = f"Fertigungsübersicht_kw{kw}.pdf"
create_pdf_with_images(arbeitsbereiche, output_filename)