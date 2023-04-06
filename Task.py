# import necessary library
import pandas as pd
import numpy as np

# download Data Frame
df = pd.read_excel('excel/Exercise.xlsx')


thickness = df['Толщина, мм'].unique().tolist()
type_of_material = df['Тип материала'].unique().tolist()
product_code = df['Код изделия'].unique().tolist()

numbers_trumpf = ['0', '4', '16', '17', '19', '22',
                  '23', '50', '54', '66', '67', '69', '72', '73']
numbers_laser = ["5", '55']

types_equipment = ['TruPunch1000', 'Laser']

# function a certain type of equipment


def type_of_equipment(row):

    if row['Толщина, мм'] == 1.5:
        equipment = types_equipment[1]
    else:
        equipment = types_equipment[0]
    return equipment

# function time distribution by equipment


def time_allocation(df):

    trumpf_t = 0
    lazer_t = 0

    i = 0
    l = 0
    ind = 0

    while ind <= len(df) - 1:

        if df['Тип оборудования'][ind] == types_equipment[0]:

            trumpf_t += df['Затрата времени вырубка (мин)'][ind]

            if trumpf_t < 100:
                df['Тип оборудования'][ind] = numbers_trumpf[i] + \
                    ' - ' + df['Тип оборудования'][ind]
                ind += 1
            else:
                trumpf_t = 0
                i += 1

        elif df['Тип оборудования'][ind] == types_equipment[1]:

            lazer_t += df['Затрата времени вырубка (мин)'][ind]

            if lazer_t < 250:
                df['Тип оборудования'][ind] = numbers_laser[l] + \
                    ' - ' + df['Тип оборудования'][ind]
                ind += 1
            else:
                lazer_t = 0
                l += 1
        else:
            break

    return df


# function concat equipment with little time
    

def concat_equip(df):

    product = df['Код изделия'].unique().tolist()

    equip = df['Тип оборудования'].unique().tolist()

    sort_equip = df[['Код изделия', 'Тип оборудования', 'Затрата времени вырубка (мин)']]\
                .groupby(['Код изделия', 'Тип оборудования'], as_index=False).sum()\
                .sort_values(['Код изделия', 'Затрата времени вырубка (мин)'], ascending=True, ignore_index=True)
    
    list_equip = sort_equip['Тип оборудования'].tolist()

    for i in product:

        for j in equip:
            
            if df.loc[(df['Код изделия'] == i) & (df['Тип оборудования'] == j), 'Затрата времени вырубка (мин)'].sum() < 40:
                
                product_equip = df.loc[(df['Код изделия'] == i) & (df['Тип оборудования'] == j)]

                group_equip = product_equip['Тип оборудования'].tolist()

                for k in range(len(types_equipment)):

                    if types_equipment[k] in group_equip:

                        df.loc[(df['Код изделия'] == i) & (df['Тип оборудования'] == j), 'Тип оборудования'] = list_equip[list_equip.index(j) + 1]
                    
    return df 

df = df.rename(
    columns={"Затарата времени вырубка (мин)": "Затрата времени вырубка (мин)"})

df['Затрата времени вырубка (мин)'] = df['Затрата времени вырубка (мин)'].str.replace(
    ',', '.').astype('float')

# a certain type of equipment
df['Тип оборудования'] = df.apply(type_of_equipment, axis=1)

# time distribution by equipment
df1 = time_allocation(df)

# concatenation equipments with less time

df2 = concat_equip(df1)

sort_equip = df2[['Код изделия', 'Тип оборудования', 'Затрата времени вырубка (мин)']]\
                 .groupby(['Код изделия', 'Тип оборудования'], as_index=False).sum()\
                 .sort_values(['Код изделия', 'Затрата времени вырубка (мин)'], ascending=True, ignore_index=True)


# list_sort_equip = sort_equip['Тип оборудования'].tolist()

# df1.to_excel('Task.xlsx', index=False)
