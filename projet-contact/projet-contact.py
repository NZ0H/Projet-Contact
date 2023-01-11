'''

    module:: projet-contact.py
   :platform: Unix, windows
   :synopsis: Permet de lancer le programme.

    moduleauthor:: Hamon Enzo <enzo.hamon@etu.univ-poitiers.fr>, Chapus Anthony <anthony.chapus@etu.univ-poitiers.fr>

'''

from fichier_excel import fichier_excel
from contact import contact
from page_web import page_web
import argparse


valeur_tmp_fct2=[]

def argument():
    """
        La fonction permet d'executer en ligne de commande avec les paramètres suivant :
            -f ou --file -> type : str , utilisation obligatoire , prend en entrer 1 ou plusieurs fichier excel.
            -s ou --sortie -> type : str, utilisation facultative , prend en entrer 1 ou plusieurs type de fichier.
            -d ou --dir -> type : str , utilisation facultative , prend en entrer un chemin de direction.
            
    :param 1: La fonction n'a pas de paramètre d'entrée.
    :returns: Renvoie les arguments ci-dessus.
    :rtype: str
    :raises: TypeError
    :example:
        Quelque Commandes
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
        valeur_tmp_fct2.append(fichier_excel(nb_arg))
    for nb_arg in arguments.sortie :
        contact(nb_arg,valeur_tmp_fct2) 
    page_web(arguments.dir,valeur_tmp_fct2)
except:
    print("Erreur durant l'utilisation du programme")