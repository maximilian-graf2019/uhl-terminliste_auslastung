import pandas as pd 
import glob

excel_files = glob.glob("*.xlsx")
file = excel_files[0]
if len(excel_files) > 1:
    raise ValueError('Es liegen mehrere Dateien im Ordner. Bitte lösche nicht mehr benötigte Dateien.')

print(f'Using file "{file}"')

# get files and read first excel into df
excel_files = glob.glob("*.xlsx")
file = excel_files[0]
df = pd.read_excel(file)

# get relevant columns and filter dataframe
cols = ['Auftrag'] + [col for col in df.columns if col.__contains__('KW')]
df_relevant = df[cols]
print(df_relevant.shape)
for col in df_relevant.columns:
    new_col = col.replace('\n', '_')
    df_relevant.rename(columns={col: new_col}, inplace=True)

# create df per 'Arbeitsbereich'
arbeitsbereiche = ['Kapazität PR-Fertigung', 'Kapazität Fensterfertigung',
                   'Kapazität Türfertigung', 'Kapazität Blechfertigung', 'Kapazität Abt. Schweißen', 'Kapazität Rollen']

row_nums = list()
for ab in arbeitsbereiche:
    rownum = df_relevant.loc[df_relevant['Auftrag'] == ab].index
    row_nums.append(rownum[0])

df_pr = df_relevant.iloc[row_nums[0]-1:row_nums[0]+1].dropna(axis=1, how='all')
df_f = df_relevant.iloc[row_nums[1]-1:row_nums[1]+1].dropna(axis=1, how='all')
df_t = df_relevant.iloc[row_nums[2]-1:row_nums[2]+1].dropna(axis=1, how='all')
df_b = df_relevant.iloc[row_nums[3]-1:row_nums[3]+1].dropna(axis=1, how='all')
df_s = df_relevant.iloc[row_nums[4]-1:row_nums[4]+1].dropna(axis=1, how='all')
df_r = df_relevant.iloc[row_nums[5]-1:row_nums[5]+1].dropna(axis=1, how='all')
