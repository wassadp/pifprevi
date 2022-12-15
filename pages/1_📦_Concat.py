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
st.title("Concat 2.0")
name_sheet_cies = "pgrm_cies"
name_sheet_af = "Programme brut"
name_sheet_oal = "affectation_oal_t2e"
st.subheader("Prévision activité AF 1 :")

uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None:
    @st.cache(suppress_st_warning=True)
    def df_af_1():
        with st.spinner('Chargemement prévision AF 1 ...'):
            df_af_1 = pd.read_excel(uploaded_file,name_sheet_af,usecols=['A/D', 'Cie Ope', 'Num Vol', 'Porteur', 'Prov Dest', 'Affectation',
                        'Service emb/deb', 'Local Date', 'Semaine', 
                        'Jour', 'Scheduled Local Time 2', 'Plage',  
                        'Pax LOC TOT', 'Pax CNT TOT', 'PAX TOT'])
        st.success("Prévision AF 1 chargée !")
        return df_af_1
    
    df_af_1 = df_af_1()      

st.subheader("Prévision activité ADP :")
uploaded_file2 = st.file_uploader("Choisir un fichier :", key=3)
if uploaded_file2 is not None:
    with st.spinner('Chargemement prévision ADP ...'):
        df_cies_1 = pd.read_excel(uploaded_file2)
    placeholder0 = st.empty()
    st.success("Prévisions chargées !")

st.subheader("Fichiers affectation :")
uploaded_file3 = st.file_uploader("Choisir le fichier affectation oal :", key=4)
if uploaded_file3 is not None:
    df_oal = pd.read_excel(uploaded_file3, name_sheet_oal)
    st.success('Affectation OAL chargée !')

name_taux = "taux affectation previ_rea"

uploaded_file4 = st.file_uploader("Choisir le fichier taux_affectation.xlsx :", key=5)
if uploaded_file3 is not None:
    df_taux = pd.read_excel(uploaded_file4, name_taux)
    st.success('Taux affectation chargée !')

    ######### Traitement #########
  
    df_af_1['Service emb/deb'] = np.where((df_af_1["A/D"]=="D") & (df_af_1["Affectation"]=="F"), 'F', df_af_1['Service emb/deb'])

    ######### Gestion des dates #########

    min_date_previ = min(df_af_1['Local Date']) # min prévi AF 1
    max_date_previ = max(df_af_1['Local Date']) # max prévi AF 2
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


    #######################################################################

    #Données pour avoir les OALs, leur affectation ainsi que le taux de 
    #       correspondance par OAL
    placeholder = st.empty()
    


    df_af_1 = df_af_1.rename(columns={"Jour":"Jour (nb)",
                                    "Service emb/deb":"Libellé terminal",
                                    "Scheduled Local Time 2":"Horaire théorique"})


    #######################################################################
    terminaux_cies = ['Terminal 2A', 
                        'Terminal 2B', 
                        'Terminal 2C', 
                        'Terminal 2D', 
                        'Terminal 3',
                        'Terminal 1']

    #        Afin d'utiliser GP et MP. Attention les prévi renseignent des MP en PP ce qui fausse pour la suite. 
    #        Piste d'amélioration


    ######### Traitement #########

    df_cies_1["Porteur"] = df_cies_1['Porteur'].str.replace({"Gros porteur":"GP",
                                                            "Moyen porteur":"MP",
                                                            "Petit porteur":"MP",
                                                            "Non renseigné":"MP"})

    data_cies_concat_1 = df_cies_1[df_cies_1["Libellé terminal"] == "Terminal 2E"]
    data_cies_concat_1 = data_cies_concat_1[data_cies_concat_1["Cie Ope"].isin(df_oal["Code IATA"].tolist()) == True]
    data_cies_concat_1.reset_index(inplace=True)
    del data_cies_concat_1['index']
    df_cies_concat_1 = df_cies_1[df_cies_1["Libellé terminal"].isin(terminaux_cies) == True]

    ######### Def #########

    placeholder.success("Mise en forme des prévisions faite !")
    #Affecter a un hall (Ici EK, EL, EM en fonction de l'oal)
    def new_DISPATCH(df, df_oal):
        list_temp = []
        df_copy = df
        k = 0
        for i in range(df_copy.shape[0]):
            for j in range(df_oal.shape[0]):
                if df_copy.iloc[i, 1] == df_oal.iloc[j, 1]:
                    
                    if df_oal.iloc[j, 2] == df_oal.iloc[j, 3]:
    #                   On met le bon libellé terminal et le taux de pax en corres
                        df_copy.loc[i, 'Libellé terminal'] = df_oal.iloc[j, 2]
                        df_copy.loc[i, 'Pax LOC TOT'] = df_copy.loc[i, 'PAX TOT'] * (1 - float(df_oal.iloc[j, 5]))
                        df_copy.loc[i, 'Pax CNT TOT'] = df_copy.loc[i, 'PAX TOT'] * float(df_oal.iloc[j, 5])
                        break
                    elif df_oal.iloc[j, 2] != df_oal.iloc[j, 3]:

                        df_copy.loc[i, 'Libellé terminal'] = df_oal.iloc[j, 2]
                        
                        df_copy.loc[i, 'Pax LOC TOT'] = df_copy.loc[i, 'PAX TOT'] * (1 - float(df_oal.iloc[j, 5])) * float(df_oal.iloc[j, 4])
                        df_copy.loc[i, 'Pax CNT TOT'] = df_copy.loc[i, 'PAX TOT'] * float(df_oal.iloc[j, 5]) * float(df_oal.iloc[j, 4])
                        df_copy.loc[i, 'PAX TOT'] = df_copy.loc[i, 'PAX TOT'] * float(df_oal.iloc[j, 4])
                        
                        list_temp.append(df_copy.loc[i].to_frame().T)
                        
                        list_temp[k]['Libellé terminal'] = df_oal.iloc[j, 3]
                        list_temp[k]['Pax LOC TOT'] = df_copy.loc[i, 'PAX TOT'] * (1 - float(df_oal.iloc[j, 5])) * (1 - float(df_oal.iloc[j, 4]))
                        list_temp[k]['Pax CNT TOT'] = df_copy.loc[i, 'PAX TOT'] * float(df_oal.iloc[j, 5]) * (1 - float(df_oal.iloc[j, 4]))
                        list_temp[k]['PAX TOT'] = list_temp[k]['PAX TOT'] * (1 - float(df_oal.iloc[j, 4]))
                        k = k + 1
                        break
                    else:
    #                            Vérifier si affectation 1 est bien dans les libellés terminaux (EK, EL, EM)
                        st.error('Erreur dans Affectation OAL : affectation 1 ou 2 invalides !', k)
        
        df_temp = pd.concat(list_temp)
        df_oal_concat = pd.concat([df_copy, df_temp])   
        df_oal_concat.reset_index(inplace=True)
        del df_oal_concat['index']    
        
        return df_oal_concat


    # pytest possible ?
    def VALID(df, eps):
        cpt = 0
        for i in range(df.shape[0]):
            if abs(df.loc[i, 'Pax LOC TOT'] + df.loc[i, 'Pax CNT TOT'] - df.loc[i, 'PAX TOT']) < eps:
    #            st.write('ligne',i,'validée')
                cpt+=1
            else:
                st.error('pas bon', df.loc[i, 'Cie Ope'], 'index :', i)
        if cpt == df.shape[0]:
            placeholder.info("Données valides")
        else:
            st.error("Erreur dans les données : PAX (LOC + CNT) <> PAX TOT")

    # A supprimer
    def CONCAT_PGRM_AF_ADP(df_af_1, df_cies_1, df_cies_oal_1):
        L = []
        L.append(df_af_1)
        L.append(df_cies_1)
    #            L.append(df_cies_2)
        L.append(df_cies_oal_1)
    #            L.append(df_cies_oal_2)
        
        df_concat = pd.concat(L)
        df_concat.reset_index(inplace=True)
        del df_concat['index']
        
        return df_concat


    data_cies_oal_concat_1 = new_DISPATCH(data_cies_concat_1, df_oal)
    placeholder.success("OAL extraites !")        


    VALID(data_cies_oal_concat_1, eps=0.1)
    #        VALID(data_cies_oal_concat_2, eps=0.1)

    ###############################################################################
    placeholder.info("Préparation à la concaténation des prévisions ...")
    placeholder.info("Récupération des champs vides ...")
    df_pgrm_concat = CONCAT_PGRM_AF_ADP(df_af_1, df_cies_concat_1, data_cies_oal_concat_1)
    x = df_pgrm_concat[df_pgrm_concat['A/D']=='D']
    y = x[x['Affectation']=='F']
    st.write(y['Pax LOC TOT'].sum())

    df_pgrm_concat['Plage'] = df_pgrm_concat['Plage'].fillna(value = "P4")


    #   A automatiser car ne prend pas toutes les cies en compte, ex ici c'est RC
    df_pgrm_concat = df_pgrm_concat.dropna(subset=['Pax LOC TOT'])

    df_pgrm_concat['Libellé terminal'].loc[(df_pgrm_concat['Cie Ope'] == 'RC')] = 'Terminal 2D'
    #df_nan['Plage'] = df_nan['Plage'].fillna(value = "P4")

    #         36% est le nomre moyen de corres pour prévision activité AF
    #df_pgrm_concat['Pax LOC TOT'] = (df_pgrm_concat['Pax LOC TOT']*(1-0.36)).astype('int')
    df_pgrm_concat.loc[(df_pgrm_concat['Pax LOC TOT'].isna()) , 'Pax LOC TOT'] = (df_pgrm_concat['Pax LOC TOT']*(1-0.36)).astype('int')
    df_pgrm_concat['Pax CNT TOT'] = 0

    # x = df_pgrm_concat[df_pgrm_concat['A/D']=='D']
    # y = x[x['Affectation']=='F']
    # st.write(y['Pax LOC TOT'].sum())

    df_pgrm_concat.loc[df_pgrm_concat['Num Vol'] == 'MNE', 'Cie Ope'] = 'ZQ'
    df_pgrm_concat.loc[df_pgrm_concat['Pax LOC TOT'] != 0, 'Pax CNT TOT'] = df_pgrm_concat['PAX TOT'] - df_pgrm_concat['Pax LOC TOT']
    
    # à ajouter : df_pgrm_concat.dropna(inplace=True)

    st.write(df_pgrm_concat['Libellé terminal'].unique())
    st.write('######"')
    st.write(df_pgrm_concat['Libellé terminal'].unique())
    placeholder.success("Concaténation des prévisions réussie !")

    ### Taux Affectation pgrm AF ###
    df_previ = df_pgrm_concat.loc[(df_pgrm_concat['Libellé terminal'].isin(['EK', 'EL', 'EM']) == True)]

    df_taux.rename(columns = {'Unnamed: 0':'Code IATA compagnie'}, inplace = True)
    df_taux = df_taux.drop(df_taux.loc[(df_taux['taux K'] == 0) & (df_taux['taux L'] == 0) & (df_taux['taux M'] == 0)].index)
    df_taux.reset_index(inplace=True)
    del df_taux['index']
            
    #On regroupe les cies dans une liste, on peut aussi changer le code pour utiliser groupby 
    l_vol_cies = []
    for cie in df_taux['Code IATA compagnie']:
        l_vol_cies.append((cie, df_previ.loc[(df_previ['Cie Ope'] == cie)]))

    l_index = []
    for l_vol_index in range(1, len(l_vol_cies)):
        l_index_l = l_vol_cies[l_vol_index][1]['Libellé terminal'].loc[l_vol_cies[l_vol_index][1]['Libellé terminal'] == 'EK'].sample(frac = df_taux['taux L'][l_vol_index]).index.tolist()
        
        for index_l in l_index_l:
            l_vol_cies[l_vol_index][1]['Libellé terminal'][index_l] = 'EL'
        
        l_index_m = l_vol_cies[l_vol_index][1]['Libellé terminal'].loc[l_vol_cies[l_vol_index][1]['Libellé terminal'] == 'EK'].sample(frac = (df_taux['taux K'][l_vol_index] + df_taux['taux M'][l_vol_index]) * df_taux['taux M'][l_vol_index]).index.tolist()
        
        for index_m in l_index_m:
            l_vol_cies[l_vol_index][1]['Libellé terminal'][index_m] = 'EM'
        
        l_index.append((l_vol_cies[l_vol_index][0], l_index_l, l_index_m))

    for index_tuple in l_index:
        df_pgrm_concat['Libellé terminal'][index_tuple[1]] = 'EL'
        df_pgrm_concat['Libellé terminal'][index_tuple[2]] = 'EM'

    ######### Export PGRM CONCAT ########      

    placeholder.info("Préparation à l'export du programme complet ...")
    directory_concat = "pgrm_complet_" + str(pd.datetime.now())[:10] + ".xlsx"
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
        writer.save()

        st.download_button(
        label="Télécharger fichier Programme complet",
        data=buffer,
        file_name=directory_concat,
        mime="application/vnd.ms-excel"
        )
    
    st.markdown('<a href="/concat_st" target="_self">Revenir à l\'Accueil</a>', unsafe_allow_html=True)
    st.markdown('<a href="/pi_previ" target="_self">Aller directement à l\'outils Pif prévi</a>', unsafe_allow_html=True)

