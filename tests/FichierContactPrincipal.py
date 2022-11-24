
import numpy as np
import pandas as pd
import openpyxl as xl
from openpyxl import Workbook,load_workbook


#./nom_projet.py --input-file fichier_annee1.xls fichier_annee2.xls  --contacts liste_contacts.vcf --output-dir ./../html
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

            for i in range(len(data)):
                data_nom_entreprise.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'Entreprise' :
                    data_nom_entreprise.remove('Entreprise')

        if column=='ADRENTR1':
            for i in range(len(data)):
                data_adresse.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'ADRENTR1' :
                    data_adresse.remove('ADRENTR1')

        if column=='LOCALITE':
            for i in range(len(data)):
                data_ville.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'LOCALITE' :
                    data_ville.remove('LOCALITE')

        if column=='CODPOST':
            for i in range(len(data)):
                data_code_postal.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'CODPOST' :
                    data_code_postal.remove('CODPOST')

        if column=='Pays':
            for i in range(len(data)):
                data_pays.append(data.iloc[i,id])
            for i in data_nom_entreprise:
                if i == 'Pays' :
                    data_pays.remove('Pays')

    return data_nom_entreprise,data_ville,data_adresse,data_code_postal


print(input('stage 2005.xlsx'))





def contact(file):
    """
    Convertir la liste des fichiers séléctionné et en faire un seul fichier au format vcf
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


















