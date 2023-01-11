'''

    module:: contact.py
   :platform: Unix, windows
   :synopsis: création d'un répertoire ou il y aura tous les fichiers vcard et xlsx.

    moduleauthor:: Hamon Enzo <enzo.hamon@etu.univ-poitiers.fr>, Chapus Anthony <anthony.chapus@etu.univ-poitiers.fr>

'''

import openpyxl
from fusion import fusion

def contact(args_sortie,valeur_tmp_fct2):
    """
    Cette fonction permet de créer un répertoire où se situera tous les fichiers vcard et xlsx

    :param args_sortie: Fichier où se situera noter répertoire.
    :type args_sortie: str
    :param valeur_tmp_fct2: Toutes les données utiles tels que Entreprise, ville , etc ...
    :type valeur_tmp_fct2: str
    :returns: Renvoie donc un répertoire avec des fichiers vcard et xlsx.
    :rtype: int
    :raises: TypeError
    :example:

    Vcard :
        BEGIN:VCARD
        VERSION:3.0
        FN:Ronan DE KERMADEC
        N:DE KERMADEC;Ronan;;;
        item1.EMAIL;TYPE=INTERNET:ronan.dekermadec@e-qual.fr
        item1.X-ABLabel:
        item2.TEL:06.17.45.12.75
        item2.X-ABLabel:
        item3.ORG:SA E-QUAL
        item3.X-ABLabel:
        NOTE:Refonte de l'infrastructure Active Directory basé sur Windows Server 2019
        CATEGORIES:myContacts
        END:VCARD

    xlsx : 
            Entreprise | Ville | Code postal | Civilite | Nom | Prénom | Sujet |       Email       | Telephone
               Orange  | Paris |    75000    |    Mr    | Bop |  Tom   | Fibre | Bop.tom@gmail.com | 06.07.08.09.10

    """

    all_value=valeur_tmp_fct2
    colonne=fusion(all_value)
    print(colonne)
    if args_sortie[len(args_sortie)-4:] == 'xlsx':
      
        wb_out = openpyxl.Workbook()
        ws1 = wb_out.active
        ws1.title = "Etudiants"
        
        ws1.append(["Nom Entreprise","Ville","Code Postal","Sujet","Civilité","Nom","Prénom","Téléphone","Email"]) 
        for entreprise,ville,code_postal,sujet,civilite,nom,prenom,tel,email in zip(colonne[0],colonne[1],colonne[2],colonne[3],colonne[4],colonne[5],colonne[6],colonne[7],colonne[8]):
            ws1.append([entreprise,ville,code_postal,sujet,civilite,nom,prenom,tel,email])
        wb_out.save('../VCARD/'+args_sortie)
    if args_sortie[len(args_sortie)-5:] == 'vcard':
        for num_c in range(len(colonne[5])):
            f= open("../VCARD/vcard_{}_{}.vcf".format(colonne[5][num_c],colonne[6][num_c]),"w")
            f.write("BEGIN:VCARD\nVERSION:3.0\n")
            f.write("FN:{} {}\n".format(colonne[6][num_c],colonne[5][num_c]))
            f.write("N:{};{};;;\n".format(colonne[5][num_c],colonne[6][num_c]))
            f.write("item1.EMAIL;TYPE=INTERNET:{}\n".format(colonne[8][num_c]))
            f.write("item1.X-ABLabel:\nitem2.TEL:{}\n".format(colonne[7][num_c]))
            f.write("item2.X-ABLabel:\nitem3.ORG:{}\n".format(colonne[0][num_c]))
            f.write("item3.X-ABLabel:\nNOTE:{}\n".format(colonne[3][num_c]))
            f.write("CATEGORIES:myContacts\nEND:VCARD")
