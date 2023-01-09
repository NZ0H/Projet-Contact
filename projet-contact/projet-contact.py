from fichier_excel import fichier_excel
from contact import contact
from fusion import fusion
from page_web import page_web
import argparse


valeur_tmp_fct2=[]

def argument():
    """
     creation : Enzo , 23/12/2022

     La fonction n'a pas de paramètre d'entrée 

     la fonction permet d'executer en ligne de commande avec les paramètres suivant :
         -f ou --file -> type : str , utilisation obligatoire , prend en entrer 1 ou plusieurs fichier excel
         -s ou --sortie -> type : str, utilisation facultative , prend en entrer 1 ou plusieurs type de fichier
         -d ou --dir -> type : str , utilisation facultative , prend en entrer un chemin de direction  

     ex d'utilisation :
        python3 projet-contact.py --file Entreprises_stage_2020.xlxs --sortie test.xlsx --dir ../html/
        python3 projet-contact.py -f Entreprises_stage_2020.xlxs -s test.xlsx -d ../html/

    """
    parser = argparse.ArgumentParser()
    parser.add_argument('-f','--file',type=str,required=True , nargs='+')
    parser.add_argument('-s','--sortie',type=str,nargs='+')
    parser.add_argument('-d','--dir',type=str)
    args=parser.parse_args()
    return args

arguments=argument()


try:
    for nb_arg in arguments.file :
        print('good1')
        valeur_tmp_fct2.append(fichier_excel(nb_arg))
        print('good2')
    for nb_arg in arguments.sortie :
        contact(nb_arg,valeur_tmp_fct2)
        print('ok')
 
    page_web(arguments.dir,valeur_tmp_fct2)
    print('good4',nb_arg)
except:
    print("Erreur durant l'utilisation du programme")