import streamlit as st
import pandas as pd

# Fonction pour récupérer les paires start_week et end_week, incluant la dernière semaine si non suivie
def get_week_intervals(df):
    week_columns = [col for col in df.columns if 'Semaine' in col]
    intervals = []
    for i in range(len(week_columns) - 1):
        start_week = week_columns[i]
        end_week = week_columns[i + 1]
        intervals.append((start_week, end_week))
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
            if end_week is None:
                if col == 'Totaux':  # Arrêter l'inclusion dès qu'on atteint 'Totaux'
                    break
                selected_columns.append(col)
            elif end_week in col:
                break
            else:
                selected_columns.append(col)
    return selected_columns

# Fonction pour obtenir les colonnes par intervalle
def get_columns_by_interval(df, week_intervals):
    columns_by_interval = {}
    for start_week, end_week in week_intervals:
        interval_columns = get_columns_between_weeks_exclusive(df, start_week, end_week)
        interval_columns.insert(0, 'Transporteur')  # Ajouter 'Transporteur' en premier
        columns_by_interval[(start_week, end_week)] = interval_columns
    return columns_by_interval

# Fonction pour créer des DataFrames par intervalle
def create_dataframes_by_interval(df, columns_by_interval):
    dataframes_by_interval = {}
    for idx, (interval, columns) in enumerate(columns_by_interval.items()):
        df_interval = df[columns].copy()
        df_interval = df_interval.iloc[:-1]
        date_S1 = "S" + df_interval.columns.to_list()[1].strip().split()[-1]
        df_interval['Date'] = date_S1
        df_interval.columns = df_interval.iloc[0]
        df_interval = df_interval[2:].reset_index(drop=True)
        df_interval.rename(columns={date_S1: 'Date'}, inplace=True)
        dataframe_name = f"S{idx + 1}"
        dataframes_by_interval[dataframe_name] = df_interval
    return dataframes_by_interval

# Fonction pour concaténer tous les DataFrames en un seul
def concatenate_dataframes(dataframes_by_interval):
    dataframes_list = list(dataframes_by_interval.values())
    df_concat = pd.concat(dataframes_list, ignore_index=True)
    return df_concat

# Fonction principale pour traiter le fichier et retourner le DataFrame final
def process_excel_file(file):
    df = pd.read_excel(file, header=11)
    week_intervals = get_week_intervals(df)
    columns_by_interval = get_columns_by_interval(df, week_intervals)
    dataframes_by_interval = create_dataframes_by_interval(df, columns_by_interval)
    df_concat = concatenate_dataframes(dataframes_by_interval)
    return df_concat

# Streamlit app
def main():
    st.title("Traitement de fichier Excel - Concatenation des DataFrames")

    # Upload du fichier Excel
    uploaded_file = st.file_uploader("Uploader un fichier Excel", type=["xlsx"])
    
    if uploaded_file is not None:
        # Affichage du fichier traité
        df_concat = process_excel_file(uploaded_file)
        st.write("DataFrame concaténé:")
        st.write(df_concat)

        # Option pour télécharger le DataFrame concaténé au format CSV
        csv = df_concat.to_csv(index=False).encode('utf-8')
        st.download_button(label="Télécharger CSV", data=csv, file_name="df_concat.csv", mime='text/csv')

if __name__ == "__main__":
    main()
