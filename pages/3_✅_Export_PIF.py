import pandas as pd
import streamlit as st
import os
import time
import pandas as pd
import xlwt
from xlwt.Workbook import *
from pandas import ExcelWriter
import xlsxwriter
import datetime
import calendar
import locale
from openpyxl.styles import Font
import itertools
from datetime import datetime, timedelta
from streamlit_extras.app_logo import add_logo
import io
from pyxlsb import open_workbook as open_xlsb
locale.setlocale(locale.LC_ALL, "fr_FR")

st.title("✅ Macro final")
#add_logo("Logo_Groupe_ADP.png")
st.write("Macro du fichier Export_pif final")

def findDay(date):
    born = datetime.datetime.strptime(date, '%d %m %Y').weekday()
    return (calendar.day_name[born])   


data = []

df_config = pd.DataFrame(data)

df_config['site'] = ['K CTR', 'K CNT', 'L CTR', 'L CNT', 'M CTR', 'Galerie EF', 'C2F', 'C2G', 'Liaison BD',
                    'T3', 'Terminal 1', 'Terminal 1_5', 'Terminal 1_6']

df_config['Abattement (%)'] = 0

 
uploaded_file = st.file_uploader("Choose a file")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df1 = pd.DataFrame(columns=df.columns)
    wb= Workbook()
    writer = pd.ExcelWriter('multiple3.xlsx', engine='xlsxwriter')

    start_date = pd.to_datetime(uploaded_file.name[14:24])
    end_date = pd.to_datetime(uploaded_file.name[28:38])  

    mask = (df['jour'] >= start_date) & (df['jour'] <= end_date)
    mask_dissocie_1 = (df['jour'] >= start_date) & (df['jour'] <= end_date - timedelta(days=7))
    mask_dissocie_2 = (df['jour'] >= start_date + timedelta(days=4)) & (df['jour'] <= end_date)

    df = df.loc[mask]
    export_pif_4_jours = df.loc[mask_dissocie_1]
    export_pif_4_jours.name = "export_pif_4jours"
    export_pif_7_jours = df.loc[mask_dissocie_2]
    export_pif_7_jours.name = "export_pif_7jours"


    st.write("Gestion de l'abattement et de l'ordre des feuilles :")

    df_config = st.data_editor(df_config)


    on = st.toggle('Dissocié')

    if not on:
        st.write('Le fichier restera unique')
        dataframe = [df]

    if on:
        st.write('Le fichier sera dissocié en deux fichiers distinct')
        dataframe = [export_pif_4_jours, export_pif_7_jours]



    def clean(df,i):
        df['Total'] = df.iloc[:, 1:145].sum(axis=1)
        df['Numéro de Jour'] = df['jour'].dt.day
        df['Date complète'] = df['jour'].dt.strftime('%d/%m/%Y')
        df['Jour de la semaine'] = df['jour'].dt.day_name(locale="fr_FR")     
        g = str(i).replace(" ", "_")
        df[str(i).replace(" ", "_")] = df['jour'].dt.month_name(locale="fr_FR")
        df["Jour férié ?"] = ""
        first_column = df.pop('Jour férié ?')
        df.insert(1, '"Jour férié ?', first_column)
        first_column = df.pop('Numéro de Jour')
        df.insert(1, 'Numéro de Jour', first_column)
        first_column = df.pop('Date complète')
        df.insert(3, 'Date complète', first_column)
        first_column = df.pop('Jour de la semaine')
        df.insert(3, 'Jour de la semaine', first_column)
        first_column = df.pop(str(i).replace(" ", "_"))
        df.insert(0, str(i).replace(" ", "_"), first_column)
        df.pop('jour')
        df[str(i).replace(" ", "_")] = list(itertools.chain.from_iterable([key] + [float('nan')]*(len(list(val))-1) 
                            for key, val in itertools.groupby(df[str(i).replace(" ", "_")].tolist())))

    
    def findDay(date):
        born = datetime.datetime.strptime(date, '%d %m %Y').weekday()
        return (calendar.day_name[born])    



  
    buffer = io.BytesIO()

    if not on:
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write each dataframe to a different worksheet.
            site = []
            for i in df_config.site.unique():
                name = str(i).replace(" ", "_")
                site += [name]
                name = df.copy()
                name = name[name['site'] == i]
                name = name.pivot_table(values='charge', index='jour', columns=['heure'], aggfunc='first')
                name.reset_index(inplace=True)
                name.fillna(0, inplace=True)
                clean(name,i)
                mask = df_config['site'] == i
                if df_config[mask]['Abattement (%)'].iloc[0] != 0:
                    for i in range(5,150):
                        name.iloc[:,i] *= (100-df_config[mask]['Abattement (%)'].iloc[0])/100

                name.to_excel(writer, sheet_name=str(i).replace(" ", "_"), index=False)
            writer.close()

            st.download_button(
            label="Télécharger fichier Export pif",
            data=buffer,
            file_name="export_pif.xlsx",
            mime="application/vnd.ms-excel"
            )

    if on:
        for df in dataframe:
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Write each dataframe to a different worksheet.
                site = []
                for i in df_config.site.unique():
                    name = str(i).replace(" ", "_")
                    site += [name]
                    name = df.copy()
                    name = name[name['site'] == i]
                    name = name.pivot_table(values='charge', index='jour', columns=['heure'], aggfunc='first')
                    name.reset_index(inplace=True)
                    name.fillna(0, inplace=True)
                    clean(name,i)
                    mask = df_config['site'] == i
                    if df_config[mask]['Abattement (%)'].iloc[0] != 0:
                        for o in range(5,150):
                            name.iloc[:,o] *= (100-df_config[mask]['Abattement (%)'].iloc[0])/100

                    name.to_excel(writer, sheet_name=str(i).replace(" ", "_"), index=False)
                writer.close()

                st.download_button(
                label="Télécharger fichier " + df.name,
                data=buffer,
                file_name= df.name + ".xlsx",
                mime="application/vnd.ms-excel"
                )
