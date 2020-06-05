import pandas as pd
# Rikoksien tasonumerointi ei ole käytössä, koska ei ainakaan vielä toimi hyvin.

sheets= [
    
    'Koko maa_2',
    'Helsinki_3',
    'Itä-Uusimaa_4',
    'Kaakkois-Suomi_5',
    'Länsi-Uusimaa_6',
    'Häme_7',
    'Sisä-Suomi_8',
    'Pohjanmaa_9',
    'Lounais-Suomi_10',
    'Itä-Suomi_11',
    'Oulu_12',
    'Lappi_13',
    
]
dfs=[]
file_name = '90029_Rikostilasto_nimikkeet-fi_2.xlsx'

for sheet in sheets:
    
    metas = []
    levels = []
    level = 1
	
    df = pd.read_excel(file_name,sheet_name = sheet)
    df.columns = df.columns.str.replace('Unnamed: 0','Alue')
    df.Alue = df.Alue.fillna(method='ffill')
    col = df.columns[1]

    prev = df[col].values[0]
    levels.append(level)
    metas.append(prev)
    for val in df[col].values[1:]:

        current = val

        if current == 'Rikokset':
            level = 1

        if prev[1:].replace('RL','rl').islower() and current[1:].replace('RL','rl').islower():

            metas.append(prev)
            
            levels.append(levels[-1]+1)

        elif prev[1:].replace('RL','rl').islower() and current[1:].replace('RL','rl').isupper():

            metas.append(prev)
            
            levels.append(levels[-1]+1)

        elif prev[1:].replace('RL','rl').isupper() and current[1:].replace('RL','rl').isupper():

            metas.append(metas[-1])
            levels.append(levels[-1])

        elif prev[1:].replace('RL','rl').isupper() and current[1:].replace('RL','rl').islower():

            metas.append(pd.unique(metas)[-2])
            
            levels.append(levels[-1]-1)
        else:
            #Debuggaus kaiken varalta. Ei pitäisi mennä tänne.
            print(prev)
            print(current)
        prev = current



    df['meta']=metas

    dfs.append(df)
rikokset = pd.concat(dfs).iloc[:,:-1]
rikokset = rikokset.set_index('Alue')
rikokset = rikokset.drop(' ',axis=1)
rikokset.to_csv('rikokset.xlsx')