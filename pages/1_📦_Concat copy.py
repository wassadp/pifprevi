import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
             
st.set_page_config(page_title="Concat", page_icon="📦", layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

######### Input #########

#   Noms des feuilles, peut changer dans le temps si qqn le modifie
st.title("📦 Concat")
name_sheet_cies = "pgrm_cies"
name_sheet_af = "Programme brut"

st.subheader("Prévision activité AF 1 :")

def get_dispatch_sat_T1(sat):
    df = pd.read_excel("fichier_config_PIF.xlsx", sheet_name="dispatch_sat")
    df = df.fillna("XXXXX")
    return list(df[sat])


uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None:
    @st.cache_data()
    def previ_af():
        with st.spinner('Chargemement prévision AF 1 ...'):
            df_af_1 = pd.read_excel(uploaded_file,name_sheet_af,usecols=['A/D', 'Cie Ope', 'Num Vol', 'Porteur', 'Type Avion', 'Prov Dest', 'Affectation',
                        'Service emb/deb', 'Local Date', 'Semaine', 
                        'Jour', 'Scheduled Local Time 2', 'Plage',  
                        'Pax LOC TOT', 'Pax CNT TOT', 'PAX TOT'])
            df_af_1.rename(columns = {'Type Avion':'Sous-type avion'}, inplace = True)
            df_af_1['Service emb/deb'] = np.where((df_af_1["A/D"]=="D") & (df_af_1["Affectation"]=="F"), 'F', df_af_1['Service emb/deb'])
            df_af_1 = df_af_1.rename(columns={"Jour":"Jour (nb)",
                                    "Service emb/deb":"Libellé terminal",
                                    "Scheduled Local Time 2":"Horaire théorique"})
        return df_af_1
    
    df_af_1 = previ_af()
    st.success("Prévision AF 1 chargée !")

    st.subheader("Prévision activité ADP :")
    uploaded_file2 = st.file_uploader("Choisir un fichier :", key=3)
    if uploaded_file2 is not None:
        def previ_adp():
            with st.spinner('Chargemement prévision ADP ...'):
                df_cies_1 = pd.read_excel(uploaded_file2)
                df_cies_1.rename(columns={"sens":"A/D",
                                "Jour":"Local Date",
                                "Nombre de passagers prévisionnels":"PAX TOT",
                                "Terminal_format_saria":"Terminal_corrigé",
                                "Numéro de vol":"Num Vol",
                                "Code IATA compagnie":"Cie Ope",
                                "Code aéroport IATA proche":"Prov Dest"},
                                inplace = True)
                df_cies_1["Pax LOC TOT"] = 0
                df_cies_1["Pax CNT TOT"] = 0
                df_cies_1 = df_cies_1.rename(columns={"Terminal_corrigé":"Libellé terminal"})
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2B","Terminal 2B")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2D","Terminal 2D")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2A","Terminal 2A")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2C","Terminal 2C")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2E","Terminal 2E")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2F","Terminal 2F")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C2G","Terminal 2G")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("C1","T1_Inter")
                df_cies_1["Libellé terminal"] = df_cies_1['Libellé terminal'].str.replace("CT","Terminal 3")
            
            return df_cies_1

        df_cies_1 = previ_adp()
        st.success("Prévisions ADP chargées !")


        st.markdown('--------------------')
        ######### Traitement #########



        ######### Gestion des dates #########

        min_date_previ = min(df_af_1['Local Date']) 
        max_date_previ = max(df_af_1['Local Date']) 
        min_date_adp = min(df_cies_1['Local Date'])
        max_date_adp = max(df_cies_1['Local Date'])

        st.warning("Plage des programmes AF/Skyteam : du " + str(min_date_previ.date()) + " au " + str(max_date_previ.date()))
        st.warning("Plage du programme ADP : du " + str(min_date_adp.date()) + " au " + str(max_date_adp.date()))

        if min_date_adp <= min_date_previ and max_date_adp >= max_date_previ:
            st.warning("Prévision d'activité est limitant")

            df_cies_1 = df_cies_1.loc[(df_cies_1['Local Date'] >= min_date_previ) & (df_cies_1['Local Date'] <= max_date_previ)]
            
        elif min_date_adp >= min_date_previ and max_date_adp <= max_date_previ:
            st.warning("Réalisé d'activité est limitant")
            
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] >= min_date_adp)]
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] <= max_date_adp)]
            
        elif min_date_adp >= min_date_previ and max_date_adp >= max_date_previ and max_date_previ >= min_date_adp:
            st.warning("Programme ADP et AF 2 limitant")
            
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] >= min_date_adp)]
            df_cies_1 = df_cies_1.loc[(df_cies_1['Local Date'] <= max_date_previ)]

        elif min_date_adp <= min_date_previ and max_date_adp <= max_date_previ and max_date_adp >= min_date_previ:
            st.warning("Programme AF 1 et ADP limitant")
            
            df_cies_1 = df_cies_1.loc[(df_cies_1['Local Date'] >= min_date_previ)]
            df_af_1 = df_af_1.loc[(df_af_1['Local Date'] <= max_date_adp)]
            
        else:
            st.warning("Les programmes AF/ADP ne se recouvrent pas, impossible de continuer"
                                    + "\n Veuillez sélectionner des programmes d'activités compatibles")

        placeholder = st.empty()

        #######################################################################
        term_adp = ["Terminal 2E", "Terminal 2G", "Terminal 2F"]
        

        ######### Traitement #########

        df_cies_1 = df_cies_1[~(df_cies_1["Libellé terminal"].isin(term_adp) == True)]

        ######### Def #########

        placeholder.success("Mise en forme des prévisions faite !")
        placeholder.info("Préparation à la concaténation des prévisions ...")
        placeholder.info("Récupération des champs vides ...")
        df_concat = pd.concat([df_af_1, df_cies_1])
        df_concat.reset_index(inplace=True)
        del df_concat['index']
        df_pgrm_concat = df_concat.copy() # inutile pour le moment
        df_pgrm_concat['Plage'] = df_pgrm_concat['Plage'].fillna(value = "P4")

        #   A automatiser car ne prend pas toutes les cies en compte, ex ici c'est RC
        df_pgrm_concat = df_pgrm_concat.dropna(subset=['Pax LOC TOT'])

        df_pgrm_concat['Libellé terminal'].loc[(df_pgrm_concat['Cie Ope'] == 'RC')] = 'Terminal 2D'
        #df_nan['Plage'] = df_nan['Plage'].fillna(value = "P4")

        #         36% est le nomre moyen de corres pour prévision activité AF
        #df_pgrm_concat['Pax LOC TOT'] = (df_pgrm_concat['Pax LOC TOT']*(1-0.36)).astype('int')
        df_pgrm_concat.loc[(df_pgrm_concat['Pax LOC TOT'].isna()) , 'Pax LOC TOT'] = (df_pgrm_concat['Pax LOC TOT']*(1-0.36)).astype('int')
        df_pgrm_concat['Pax CNT TOT'] = 0

        df_pgrm_concat.loc[df_pgrm_concat['Num Vol'] == 'MNE', 'Cie Ope'] = 'ZQ'
        df_pgrm_concat.loc[df_pgrm_concat['Pax LOC TOT'] != 0, 'Pax CNT TOT'] = df_pgrm_concat['PAX TOT'] - df_pgrm_concat['Pax LOC TOT']
        sat5 = get_dispatch_sat_T1("sat5")
        sat6 = get_dispatch_sat_T1("sat6")
        df_pgrm_concat.loc[df_pgrm_concat['Cie Ope'].isin(sat6), 'Libellé terminal'] = 'Terminal 1_6'
        df_pgrm_concat.loc[df_pgrm_concat['Cie Ope'].isin(sat5), 'Libellé terminal'] = 'Terminal 1_5'
        # à ajouter : df_pgrm_concat.dropna(inplace=True)
        placeholder.success("Concaténation des prévisions réussie !")

        ######### Export PGRM CONCAT ########      
        from datetime import datetime
        placeholder.info("Préparation à l'export du programme complet ...")
        directory_concat = "pgrm_complet_" + str(datetime.now())[:10] + ".xlsx"
        df_pgrm_concat.to_excel(directory_concat, sheet_name = "pgrm_complet")
        placeholder.success("Programme complet exporté !")
        placeholder.info("Fin du traitement")
        
        import io
        from pyxlsb import open_workbook as open_xlsb

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write each dataframe to a different worksheet.
            df_pgrm_concat.to_excel(writer, sheet_name= "pgrm_complet")
            # Close the Pandas Excel writer and output the Excel file to the buffer
            writer.close()

            st.download_button(
            label="Télécharger fichier Programme complet",
            data=buffer,
            file_name=directory_concat,
            mime="application/vnd.ms-excel"
            )
        

        st.stop()
        # st.markdown('<a href="/" target="_self">Revenir à l\'Accueil</a>', unsafe_allow_html=True)
        # st.markdown('<a href="/Pif_Previ_" target="_self">Aller directement à l\'outils Pif prévi</a>', unsafe_allow_html=True)

