from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
import glob
import os
import seaborn as sns
import matplotlib.pyplot as plt
import logging

def get_techniker(filename):
    df_techniker = pd.read_excel(filename, sheet_name='Grafiken TB', skiprows=2, usecols=['Name', 'Std / Woche'])
    logging.info(df_techniker.head())
    techniker = df_techniker['Name'].tolist()
    caps = df_techniker['Std / Woche'].tolist()
    return techniker, caps

logging.basicConfig(level=logging.INFO)

excel_files = glob.glob("*.xlsx")
file = excel_files[0]
if len(excel_files) > 1:
    raise ValueError('Es liegen mehrere Dateien im Ordner. Bitte lösche nicht mehr benötigte Dateien.')

logging.info(f'Nutze die Datei "{file}"')

# getting the techniker and caps from the excel file
techniker, caps = get_techniker(file)

# modify techniker to get the correct indexes
techniker.append("Fremdleistung TB")
caps.append(0)
print(f"Gefundene Techniker: {techniker}, mit Kapazitäten: {caps}")


# remove first elements from techniker and caps (there are words instead of numbers)
# techniker, caps = techniker[1:], caps[1:]

# replacing Urlaub with 0 in caps
caps = [0 if c == 'Urlaub' else c for c in caps]


# get files and read first excel into df
excel_files = glob.glob("*.xlsx")
file = excel_files[0]
df = pd.read_excel(file)
logging.info('Excel erfolgreich eingelesen.')

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
kw_cols = [col for col in df.columns if col.__contains__('KW')]
cols = ['Bezeichnung'] + kw_cols
df_relevant = df[cols]
df_relevant[kw_cols] = df_relevant[kw_cols].apply(pd.to_numeric, errors='coerce', axis=1)
logging.info(df_relevant.shape)
for col in df_relevant.columns:
    new_col = col.replace('\n', '_')
    df_relevant.rename(columns={col: new_col}, inplace=True)
logging.info("DataFrame mit relevanten Spalten erstellt.\n")
print(df_relevant.head())

# creating variables
today = datetime.today()
kw = today.isocalendar().week

# logging.infoing the used Variables
logging.info(f'Es werden die Grafiken für die KW {kw} erstellt.')
logging.info('Folgende Techniker mit maximalen Kapazitäten werden verwendet:')
for i, t in enumerate(techniker):
    print(i, '\t', caps[i], '\t', t)

# create df per 'Techniker'
row_nums = list()
for t in techniker:
    rownum = df_relevant.loc[df_relevant['Bezeichnung'] == t].index
    logging.info(rownum)
    row_nums.append(rownum[0])

# remove the first rownum, because the values for each techniker are at the start of the next technikers row
row_nums = row_nums[1:]

# remove last element from caps and techniker
caps = caps[:-1]
techniker = techniker[:-1]

# ensuring, row_nums is correct length
assert len(row_nums) == len(caps)
assert len(row_nums) == len(techniker)

logging.info(f'Gefunden Startzeilen {row_nums}\n')
logging.info(df_relevant.head())

dfs = []

for i, t in enumerate(techniker):
    df_temp = df_relevant.iloc[row_nums[i]-1:row_nums[i] +
                         1].dropna(axis=1, how='all').transpose().stack().reset_index()[1:]
    df_temp.rename(columns={'level_0': 'KW',
                'level_1': 'Legende', 0: 'value'}, inplace=True)
    # logging.info(df_temp.info())
    row_Auslastung = row_nums[i]-1
    row_Kapazitaet = row_nums[i]
    df_temp['Legende'].replace(row_Auslastung, 'Auslastung', inplace=True)
    df_temp['Legende'].replace(row_Kapazitaet, 'Kapazität', inplace=True)
    dfs.append(df_temp) 

logging.info('Grafiken wie benötigt erstellt.\n')

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
def plot_abteilung(plot_item, data, capacity, kw=today.isocalendar().week):
    sns.set(rc={'figure.figsize': (10, 6)}, font_scale = 1.2)
    sns.catplot(data=data,
                kind='bar',
                x='KW',
                y='value',
                hue='Legende',
                height=6,
                aspect=2.5,
                hue_order=['Kapazität', 'Auslastung'],
                palette=sns.color_palette(['green', 'red']))
    plt.ylabel('Stunden')
    plt.xticks(rotation=45)
    plt.xlabel('Kalenderwochen')
    plt.axhline(capacity, c='gray')
    plt.savefig(
        f'./grafiken/Grafik_{plot_item}_KW{kw}.png', 
        bbox_inches="tight")

# get total auslastung
def get_total_auslastung(dfs):
    totals = list()
    for df in dfs:
        df_capacity = df.loc[(df.Legende == "Auslastung") & (df.KW.isin(get_kw_names(26)))].reset_index(drop=True)
        sum_capacity = float(df_capacity['value'].sum())
        totals.append(round(sum_capacity, 2))
    return {t : totals[i] for i, t in enumerate(techniker)}

totals = get_total_auslastung(dfs)
logging.info(totals)

# plotting and saving using the functions
for df_nr, t in enumerate(techniker):
    df_current = dfs[df_nr]
    plot_abteilung(
        plot_item=t,
        data=df_current.loc[df_current['KW'].isin(get_kw_names(26))],
        capacity=caps[df_nr])

logging.info('PDF wird erzeugt.\n')


def create_pdf_with_images(techniker_list, output_filename):
    doc = SimpleDocTemplate(output_filename, pagesize=A4)

    # Prepare the list of elements for the PDF
    elements = []
    styles = getSampleStyleSheet()

    # Add title to the PDF
    title = "Auslastung der Techniker"
    title_text = Paragraph(title, styles['Title'])
    elements.append(title_text)
    elements.append(Spacer(1, 20))

    for i, techniker in enumerate(techniker_list):
        # Create a heading for each techniker
        heading_text = f"Techniker: {techniker}"
        heading = Paragraph(heading_text, styles['Heading2'])
        elements.append(heading)

        # Find and add the corresponding plot image to the PDF
        image_filename = f"grafiken/Grafik_{techniker}_KW{kw}.png"
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
output_filename = f"auslastung_techniker_kw{kw}.pdf"
create_pdf_with_images(techniker, output_filename)