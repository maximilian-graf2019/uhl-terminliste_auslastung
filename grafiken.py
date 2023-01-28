from datetime import datetime
from fpdf import FPDF
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

# removing old files from subfolder grafiken 
filenames = os.listdir('grafiken')
for filename in filenames:
    os.unlink(f'grafiken/{filename}')

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
                'level_1': 'typ', 0: 'value'}, inplace=True)
    # print(df_temp.info())
    row_Auslastung = row_nums[i]-1
    row_Kapazitaet = row_nums[i]
    df_temp['typ'].replace(row_Auslastung, 'Auslastung', inplace=True)
    df_temp['typ'].replace(row_Kapazitaet, 'Kapazität', inplace=True)
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
            kws = [str(year) + '_KW0' + str(i) for i in range(1, 10)] + [str(year) + '_KW' + str(i) for i in range(10, number_of_keys)]
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
                hue='typ',
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

# plotting and saving using the functions
for df_nr, abt in enumerate(arbeitsbereiche):
    df_current = dfs[df_nr]
    plot_abteilung(
        abteilung=abt,
        data=df_current.loc[df_current['KW'].isin(get_kw_names(26))],
        capacity=caps[df_nr])

print('PDF wird erzeugt.')

title = 'Übersicht Fertigung Auslastung + Kapazität'

class PDF(FPDF):
    def header(self):
        # Arial bold 15
        self.set_font('Arial', 'B', 16)
        # # Calculate width of title and position
        # w = self.get_string_width(title) + 6
        # self.set_x((200 - w) / 2)
        # # Title
        # self.cell(w, 25, title, 0, 1, align='C')
        # Line break
        self.ln(5)

    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        # Arial italic 8
        self.set_font('Arial', 'I', 8)
        # Text color in gray
        self.set_text_color(128)
        # Page number
        self.cell(0, 10, 'Seite ' + str(self.page_no()), 0, 0, 'C')

pdf = PDF(orientation='L')
pdf.set_title(title)
pdf.add_page()
pdf.cell(0, 5, f'Pfosten-Riegel Fertigung', align='L')
pdf.image(f'grafiken/Grafik_PR-Fertigung_KW{kw}.png', x=50, y=25, w=200, h=72)
pdf.ln(85)
pdf.cell(0, 5, f'Fenster Fertigung', align='L')
pdf.image(f'grafiken/Grafik_Fensterfertigung_KW{kw}.png', x=50, y=115, w=200, h=72)
pdf.add_page()
pdf.cell(0, 5, f'Blech Fertigung', align='L')
pdf.image(f'grafiken/Grafik_Blechfertigung_KW{kw}.png', x=50, y=25, w=200, h=72)
pdf.ln(85)
pdf.cell(0, 5, f'Schweissen Fertigung', align='L')
pdf.image(f'grafiken/Grafik_Schweissen_KW{kw}.png', x=50, y=115, w=200, h=72)
pdf.add_page()
pdf.cell(0, 5, f'Türen Fertigung', align='L')
pdf.image(f'grafiken/Grafik_Türfertigung_KW{kw}.png', x=50, y=25, w=200, h=72)
pdf.ln(85)
pdf.cell(0, 5, f'Rollen Fertigung', align='L')
pdf.image(f'grafiken/Grafik_Rollen_KW{kw}.png', x=50, y=115, w=200, h=72)
    
pdf.set_author('Maximilian Graf')
pdf.output(f'Fertigungsübersicht_{today.isocalendar().week}.pdf', 'F')
