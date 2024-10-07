import pandas as pd
import numpy as np

# Fonction pour récupérer les paires start_week et end_week, incluant la dernière semaine si non suivie
def get_week_intervals(df):
    week_columns = [col for col in df.columns if 'Semaine' in col]
    intervals = []
    for i in range(len(week_columns) - 1):
        start_week = week_columns[i]
        end_week = week_columns[i + 1]
        intervals.append((start_week, end_week))
    # Ajouter la dernière semaine seule si elle n'a pas d'intervalle suivant
    if len(week_columns) > 0:
        intervals.append((week_columns[-1], None))
    return intervals

# Fonction pour récupérer les colonnes entre deux semaines, incluant le cas où end_week est None
def get_columns_between_weeks_exclusive(df, start_week, end_week):
    selected_columns = []
    include = False
    for col in df.columns:
        if start_week in col:
            include = True
        if include:
            # Si end_week est None, inclure toutes les colonnes restantes sauf 'Totaux' et suivantes
            if end_week is None:
                if col == 'Totaux':  # Arrêter l'inclusion dès qu'on atteint 'Totaux'
                    break
                selected_columns.append(col)
            # Sinon, arrêter l'inclusion avant d'atteindre end_week
            elif end_week in col:
                break
            else:
                selected_columns.append(col)
    return selected_columns

# Fonction principale pour lire le fichier et transformer les données
def process_excel_file(filename):
    df = pd.read_excel(filename, header=11)

    # Récupérer les paires de semaines
    week_intervals = get_week_intervals(df)

    # Initialiser un dictionnaire pour stocker les colonnes par intervalle
    columns_by_interval = get_columns_by_interval(df, week_intervals)

    # Dictionnaire pour stocker les DataFrames par intervalle
    dataframes_by_interval = create_dataframes_by_interval(df, columns_by_interval)

    # Concaténer les DataFrames
    df_concat = concatenate_dataframes(dataframes_by_interval)

    return df_concat

# Fonction pour obtenir les colonnes par intervalle
def get_columns_by_interval(df, week_intervals):
    columns_by_interval = {}
    for start_week, end_week in week_intervals:
        # Récupérer les colonnes pour cet intervalle spécifique
        interval_columns = get_columns_between_weeks_exclusive(df, start_week, end_week)
        # Ajouter la colonne 'Transporteur' en premier
        interval_columns.insert(0, 'Transporteur')
        # Stocker le résultat dans le dictionnaire avec l'intervalle comme clé
        columns_by_interval[(start_week, end_week)] = interval_columns
    return columns_by_interval

# Fonction pour créer des DataFrames par intervalle
def create_dataframes_by_interval(df, columns_by_interval):
    dataframes_by_interval = {}
    for idx, (interval, columns) in enumerate(columns_by_interval.items()):
        # Créer un DataFrame pour chaque intervalle
        df_interval = df[columns].copy()
        df_interval = df_interval.iloc[:-1]
        date_S1 = "S" + df_interval.columns.to_list()[1].strip().split()[-1]

        # Ajouter et transformer les colonnes
        df_interval['Date'] = date_S1
        df_interval.columns = df_interval.iloc[0]
        df_interval = df_interval[2:].reset_index(drop=True)
        df_interval.rename(columns={date_S1: 'Date'}, inplace=True)

        # Nommer le DataFrame (S1, S2, etc.)
        dataframe_name = f"S{idx + 1}"
        dataframes_by_interval[dataframe_name] = df_interval
    return dataframes_by_interval

# Fonction pour concaténer tous les DataFrames en un seul
def concatenate_dataframes(dataframes_by_interval):
    dataframes_list = list(dataframes_by_interval.values())
    df_concat = pd.concat(dataframes_list, ignore_index=True)
    return df_concat

# Appel de la fonction principale
filename = "./Taux de régularité DSP - 18092024131617.xlsx"
df_concat = process_excel_file(filename)

# Afficher le DataFrame concaténé
print(df_concat.shape)
print(df_concat.head(10))
