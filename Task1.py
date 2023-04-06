import pandas as pd
import numpy as np
#### Download DataFrame
df = pd.read_excel("excel/1501112.xls", sheet_name="Общая")

df_punch = pd.read_excel("excel/1501112.xls", sheet_name="Оборуд.", header=None)

df_bend = pd.read_excel("excel/1501112.xls", sheet_name='Оборуд гибка', header=None)
#### Clean data frame
df = df.dropna(how="all")

df["Код изделия"] = df["Код изделия"].fillna(method='ffill', axis=0)

df = df.reset_index(drop=True)

df = df.drop(columns=[
    'Код 1С', 'Задание', 'Заимствование', 
    'Количество по заданию.1', 'материнское задание', 'Операции',
    'Норма\nВырубка', 'Норма\nГибка' 
    ])

df["Тип материала"] = df["Тип материала"].str.strip()
df_punch = df_punch.drop(columns=[1])

df_punch = df_punch[ ~df_punch[0].str.contains("Amada") ]

df_punch
#### Numbers and type of equipments
df_punch = df_punch[0].str.split(pat="-", expand=True)

df_punch = df_punch.rename({0: "numbers", 1: "type_equip"}, axis=1)
    
df_punch = df_punch.sort_values(by='numbers', key=lambda x: df_punch['numbers'].str[1]).reset_index(drop=True)

df_punch
df_bend = df_bend[0].str.split(pat="-", expand=True)

df_bend = df_bend.rename({0: "numbers", 1: "type_equip"}, axis=1)

df_bend
### Total time for each product
df.groupby("Код изделия", as_index=False)[["Затраты времени вырубка"]].sum()
#### Find fuzzy duplicates
df['fuzzy_duplicates'] = df['Код изделия'].str.split('[0-9]').str.get(0)

df.head()
#### Function to determine the type of TruPunch equipment for the task
def type_trupunch(row):
    
    if row["Толщина, мм"] <= 1.2:

        equip = "TruPunch 1000"
    
    else:

        equip = "TruLaser 5030 fiber" 
    
    return equip

df["Тип оборудования 1"] = df.apply(type_trupunch, axis=1)

df.head()

#### Find cumulative sum and sort by total time product
df['cumsum'] = df.groupby(['fuzzy_duplicates', 'Тип материала', 'Толщина, мм'])['Затраты времени вырубка'].cumsum()

df['total_time'] = df.groupby(['fuzzy_duplicates', 'Тип оборудования 1'])['Затраты времени вырубка'].transform('sum')

# df['total_material'] = df.groupby(['Код изделия', 'Тип материала'])['Затраты времени вырубка'].transform('sum')

df = df.sort_values(by=['total_time', 'Толщина, мм', 'Тип материала', 'ВНС'], ignore_index=True, ascending=False)

df.head()
#### Function time distribution by equipment TruPunch
def time_distribution(df, type_equip, time=700):

    product = list(df.loc[df['Тип оборудования 1'] == type_equip, 'fuzzy_duplicates'].unique())

    numbers = list(df_punch.loc[df_punch['type_equip'] == type_equip, 'numbers'])

    k = 0

    res = 0

    i = 0

    f4 = (df['Тип оборудования 1'] == type_equip)

    while i <= len(product) - 1:

        f = (df['fuzzy_duplicates'].isin([product[i]]))

        total = df.loc[f4 & f, 'Затраты времени вырубка'].sum()

        res += total

        if res <= time:

            df.loc[f4 & f, 'numbers'] = numbers[k]

            i += 1

        elif total >= time:

            res_t = res

            material = list(df.loc[f4 & f, 'Тип материала'].unique())

            for m in material:

                f1 = (df['Тип материала'] == m)

                thinkers = list(df.loc[f4 & f & f1, "Толщина, мм"].unique())

                t = 0

                while t <= len(thinkers) - 1:

                    f2 = (df['Толщина, мм'] == thinkers[t])

                    total_t = df.loc[f4 & f & f1 & f2, 'Затраты времени вырубка'].sum()

                    if total_t >= time:

                        a = 0

                        b = total_t / time

                        c = total_t / b

                        while a <= total_t:

                            f3 = (df.loc[f4 & f & f1 & f2, "cumsum"].between(a, c))

                            df.loc[f4 & f & f1 & f2 & f3, 'numbers'] = numbers[k]

                            a += c

                            c += c

                            k += 1

                        else:

                            t += 1

                    else:

                        res_t += total_t

                        if res_t <= time:

                            df.loc[f4 & f & f1 & f2, "numbers"] = numbers[k]

                            t += 1

                        else:

                            res_t = 0

                            k += 1

            res = 0

            if res >= 500:

                res = 0

            res_t = 0

            i += 1

            k += 1

        else:

            k += 1

            res = 0

    return df
def less_time(df, type_equip):

    f = (df['Тип оборудования 1'] == type_equip)

    group = df[ f ].groupby(['numbers'], as_index=False)[['Затраты времени вырубка']].sum() 

    f = group.loc[ group['Затраты времени вырубка'] < 400,  'numbers']

    numbers = list(f)

    res = 0

    i = 0

    for n in numbers:

        f = (df['numbers'] == n)

        total = df.loc[ f, 'Затраты времени вырубка' ].sum()

        res += total

        if res <= 700:

            df.loc[ f, 'numbers' ] = numbers[i]
        
        else:

            res = 0

            i += 1
    
    return df
def time_equip(df, list_equip, fun):

    for e in list_equip:

        type_equip = e

        fun(df, type_equip)

    return(df)
list_equip = list(df['Тип оборудования 1'].unique())

df2 = time_equip(df, list_equip, fun=time_distribution)

df2.groupby(["numbers"], as_index=False)[["Затраты времени вырубка"]].sum()

df3 = time_equip(df2, list_equip, fun=less_time)

df3.groupby(["numbers"], as_index=False)[["Затраты времени вырубка"]].sum()
#### Function time distribution by equipment TruBend
def type_trubend(row):

    type_bend=list(df_bend['type_equip'].unique())
         
    if row['Количество по заданию'] >= 100 and 'Amada APX' in type_bend:

            equip = 'Amada APX'

    elif row['Тип материала'] == 'Алюмоцинк':

        equip = 'Salgvanini'

    else:

        equip = 'Trumpf' 
    
    return equip
df3['Тип оборудования гибка'] = df3.apply(type_trubend, axis=1)

df3.head()
df3['cumsum_bend'] = df3.groupby(['numbers', 'Тип оборудования гибка', 'Тип материала', 'Толщина, мм'])['Затраты времени гибка'].cumsum()
df3['total_bend'] = df3.groupby(['numbers', 'Тип оборудования гибка'])['Затраты времени гибка'].transform('sum')
df3 = df3.sort_values('total_bend', ascending=False)

df3
def time_distribution_bend(df, type_equip, time=450):

    numbers = list(df['numbers'].unique())

    numbers_bend = list(df_bend.loc[df_bend['type_equip'] == type_equip, 'numbers'])

    k = 0

    res = 0

    res_t = 0

    i = 0

    f4 = (df['Тип оборудования гибка'] == type_equip)

    while i <= len(numbers) - 1:

        f = (df['numbers'].isin([numbers[i]]))

        total = df.loc[f4 & f, 'Затраты времени гибка'].sum()

        res += total

        if res <= time:

            df.loc[f4 & f, 'numbers bend'] = numbers_bend[k]

            i += 1

        elif total >= time:
            
            material = list(df.loc[f4 & f, 'Тип материала'].unique())

            for m in material:

                f1 = (df['Тип материала'] == m)

                thinkers = list(df.loc[f4 & f & f1, "Толщина, мм"].unique())

                t = 0

                while t <= len(thinkers) - 1:

                    f2 = (df['Толщина, мм'] == thinkers[t])
                   
                    total_t = df.loc[f4 & f & f1 & f2, 'Затраты времени гибка'].sum()

                    if total_t >= time:

                        a = 0

                        b = total_t / time

                        c = total_t / b

                        while a <= total_t:

                            f3 = (df.loc[f4 & f & f1 & f2, "cumsum_bend"].between(a, c))

                            df.loc[f4 & f & f1 & f2 & f3, 'numbers bend'] = numbers_bend[k]

                            a += c

                            c += c

                            k += 1

                        else:

                            t += 1

                    else:

                        res_t += total_t

                        if res_t <= time:

                            df.loc[f4 & f & f1 & f2, "numbers bend"] = numbers_bend[k]

                            t += 1

                        else:

                            res_t = 0

                            k += 1

            res = 0

            if res >= 250:

                res = 0

            res_t = 0

            i += 1

            k += 1

        else:

            k += 1

            res = 0

    return df
list_bend = list(df3['Тип оборудования гибка'].unique())

df4 = time_equip(df3, list_equip=list_bend, fun=time_distribution_bend)

df4.head()

df4.groupby(['numbers bend', 'Тип оборудования гибка'])[['Затраты времени гибка']].sum()
def less_time(df, type_equip):

    f = (df['Тип оборудования гибка'] == type_equip)

    group = df[ f ].groupby(['numbers bend'], as_index=False)[['Затраты времени гибка']].sum() 

    f = group.loc[ group['Затраты времени гибка'] < 300,  'numbers bend']

    numbers = list(f)

    res = 0

    i = 0

    for n in numbers:

        f = (df['numbers bend'] == n)

        total = df.loc[ f, 'Затраты времени гибка' ].sum()

        res += total

        if res <= 450:

            df.loc[ f, 'numbers bend' ] = numbers[i]
        
        else:

            res = 0

            i += 1
    
    return df
df5 = time_equip(df4, list_equip=list_bend, fun=less_time)

df5.groupby(['numbers bend', 'Тип оборудования гибка'])[['Затраты времени гибка']].sum()
df5.head()
df5['Тип оборудования'] = df5['numbers'] + ' - ' + df5['Тип оборудования 1']

df5['Тип оборудования гибка'] = df5['numbers bend'] + ' - ' + df5['Тип оборудования гибка']
 
df5.head()
df5 = df5.drop(columns=['fuzzy_duplicates', 'cumsum', 'total_time', 'cumsum_bend', 'total_bend', 'Тип оборудования 1', 'numbers', 'numbers bend'])

df5
grouped_df = df5.groupby('Тип оборудования')
   
for data in grouped_df:
                       
    grouped_df.get_group(data[0]).to_excel("TruPunch/" + data[0] + ".xlsx", index=False)
grouped_df = df5.groupby('Тип оборудования гибка')

for data in grouped_df:

    grouped_df.get_group(data[0]).to_excel("TruBend/" + data[0] + ".xlsx", index=False)