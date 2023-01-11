'''

    module:: page_web.py
   :platform: Unix, windows
   :synopsis: création de la page html du projet.

    moduleauthor:: Hamon Enzo <enzo.hamon@etu.univ-poitiers.fr>, Chapus Anthony <anthony.chapus@etu.univ-poitiers.fr>

'''

import os 
from fusion import fusion

def page_web(args,valeur_tmp_fct2):  

    '''
    Cette fonction génère une page html sous forme d'un tableau avec toutes les données utiles.

    :param args: Fichier où se situera noter index.html.
    :type args: str
    :param valeur_tmp_fct2: Toutes les données utiles tels que Entreprise, ville , etc ...
    :type valeur_tmp_fct2: str
    :returns: Renvoie donc un fichier html.
    :rtype: int
    :raises: TypeError
    :example:

        Entreprise | Ville | Code postal | Civilite | Nom | Prénom | Sujet |       Email       | Telephone
           Orange  | Paris |    75000    |    Mr    | Bop |  Tom   | Fibre | Bop.tom@gmail.com | 06.07.08.09.10
    '''

    structure = ["Entreprise","Ville","Code postal","Civilite", "Nom", "Prénom","Sujet", "Email", "Telephone"]
    data = fusion(valeur_tmp_fct2)

    f = open(args+'/index.html',"w")
    f.write("<!DOCTYPE html>\n")
    f.write("<head> \n<meta charset='utf-8'>\n")
    f.write("<link rel='stylesheet' href='style.css'>")
    f.write("\n<title> Tableau </title> \n</head>")
    f.write("\n<body>\n")
    f.write("<table>\n<tr>")

    for titre in structure:
        f.write("<th>{}</th>".format(titre))
    f.write("</tr>") 

    for donnee in range(len(data[0])):
        f.write("<tr>")
        for cell in [data[0][donnee], data[1][donnee],data[2][donnee],data[4][donnee], data[5][donnee], data[6][donnee], data[3][donnee],data[8][donnee], data[7][donnee]]:
            f.write(f"<td>{cell}</td>")
        f.write("</tr>")

    f.write("</table>\n</body>\n</html>")
    f.close()
