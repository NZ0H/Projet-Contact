from fichier_excel import fichier_excel
from type_fichier import contact
from fusion import fusion
from page_web import page_web

import openpyxl
import argparse
import os

valeur_tmp_fct2=[]

def argument():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f','--file',type=str,required=True , nargs='+')
    parser.add_argument('-s','--sortie',type=str,nargs='+')
    parser.add_argument('-d','--dir',type=str,nargs='+')
    args=parser.parse_args()
   
    return args

arguments=argument()




try:
    for nb_arg in arguments.file :
        print("debut")
        print(nb_arg)
        valeur_tmp_fct2.append(fichier_excel(nb_arg))
        print("fin")
  
    for nb_arg in arguments.sortie :
        print("Début")
        contact(nb_arg,valeur_tmp_fct2)
        print("fin")
    
    for nb_arg in arguments.dir :
        page_web(nb_arg,valeur_tmp_fct2)
    print("finis")
except:
    print("Erreur durant l'utilisation du programme")