"""IMPORTATION DES BIBLIOTHEQUE UTILE AU FONCTION DU PROGRAMME"""

import numpy as np
import pandas as pd
import openpyxl as xl
from openpyxl import Workbook,load_workbook

""" forme de l'execution en ligne de commande linux"""
#./nom_projet.py --input-file fichier_annee1.xls fichier_annee2.xls  --contacts liste_contacts.vcf --output-dir ./../html

"""TOUTE LES INFORMATION DES TABLEAUX SERONT STOCKER DANS CHAQUES CATHEGORIE ADEQUAT"""

data_siret=[]
data_nom_entreprise=[]
data_ville=[]
data_adresse=[]
data_code_postal=[]
data_pays=[]



""" CREER UNE VARIABLE QUI REMPLIE UN TABLEAU SI IL EST VIDE, IL DOIT AVOIR LA MEME LONGUEUR UE LES AUTRES  """
def input(file):
    data=pd.read_excel(file,engine='openpyxl')

    for id,column in enumerate(data.columns):

        if column=='Entreprise':

            for i in range(len(data)):                  # Si le nom de la colonne égal Entreprise alors on parcourt ligne Entreprise et on ajoute chaque Entreprise dans la liste data_nom_entreprise
                data_nom_entreprise.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'Entreprise' :                   # Si le nom 'Entreprise' est présent dans la liste alors tout les mots appelés 'Entreprise' seront supprimer
                    data_nom_entreprise.remove('Entreprise')

        if column=='ADRENTR1':                          # Si le nom de la colonne égal ADRENTR1 alors on parcourt ligne ADRENTR1 et on ajoute chaque adrsse d'entreprise dans la liste data_adresse
            for i in range(len(data)):
                data_adresse.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'ADRENTR1' :                     # Si le nom ADRENTR1 est présent dans la liste alors tout les mots appelés 'ADRENTR1' seront supprimer
                    data_adresse.remove('ADRENTR1')

        if column=='LOCALITE':                          # Si le nom de la colonne égal LOCALITE alors on parcourt ligne LOCALITE et on ajoute chaque ville dans la liste data_ville
            for i in range(len(data)):
                data_ville.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'LOCALITE' :                     # Si le nom LOCALITE est présent dans la liste alors tout les mots appelés 'LOCALITE' seront supprimer
                    data_ville.remove('LOCALITE')

        if column=='CODPOST':                           # Si le nom de la colonne égal CODPOST alors on parcourt ligne CODPOST et on ajoute chaque code postal dans la liste data_code_postal
            for i in range(len(data)):
                data_code_postal.append(data.iloc[i,id])
            for i in data_nom_entreprise:                # Si le nom 'CODPOST' est présent dans la liste alors tout les mots appelés 'CODPOST' seront supprimer
                if i == 'CODPOST' :
                    data_code_postal.remove('CODPOST')

        if column=='Pays':                             # Si le nom de la colonne égal Pays alors on parcourt ligne Pays et on ajoute chaque pays dans la liste data_pays
            for i in range(len(data)):
                data_pays.append(data.iloc[i,id])
            for i in data_nom_entreprise:               # Si le nom 'Pays' est présent dans la liste alors tout les mots appelés 'Pays' seront supprimer
                if i == 'Pays' :
                    data_pays.remove('Pays')

    return data_nom_entreprise,data_ville,data_adresse,data_code_postal


print(input('stage 2005.xlsx'))




"""supprimer les chiffres générés dans la colonne excel"""

def contact(file):
    """
    Convertir la liste des fichiers séléctionné et en faire un seul fichier au format vcf

    Toute les collones à faire dans le fichiers .xlsx
    colone 1 : siret
    colone 2 : nom_entreprise
    colone 3 : adresse
    colone 4 : ville
    colone 5 : code_postal
    colone 6 : pays
    """


    all_value = input (file)
    print(len(all_value[0]))
    print(len(all_value[1]))
    print(len(all_value[2]))
    print(len(all_value[3]))



    print(all_value[0])
    wb=Workbook()
    ws = wb.active
    ws.title = 'Test'
    wb.save(file)

    donnee=pd.DataFrame({"nom_entreprise":all_value[0]})
    datatoexcel= pd.ExcelWriter(file)
    donnee.to_excel(datatoexcel)
    datatoexcel.save()


contact('test2.xlsx')



def output(dir):
    """"
    Création du tableau dans le fichier html
    """
    pass


















