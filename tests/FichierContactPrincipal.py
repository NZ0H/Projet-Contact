import openpyxl
import argparse

#Fichier pour tester la suppression des lignes None

def fichier_excel(file):
    workbook = openpyxl.load_workbook(file, data_only = True)
    titres_onglets = workbook.sheetnames
    onglet = workbook[titres_onglets[0]]

    nb_none=0

    listes_data_depart = [] #stocke toute les lignes
    listes_data_traiter = []

    data_nom_entreprise=[]
    data_lieu=[]
    data_sujet=[]
    data_tuteur=[]
    data_tel=[]
    data_mail=[]
    data_code_postal=[]
    data_prenom_tuteur=['Prenom']
    data_nom_tuteur=['Nom']
    lettre_maj=0

    for row in onglet.values: #ajoute les lignes dans la liste lignes

        listes_data_depart.append(list(row))

    for r in range(len(listes_data_depart)):        #parcourt les listes présente dans lignes

        for non in listes_data_depart[r]:       #parcourt liste de liste pour avoir les valeurs individuelles
            if non == None:
                nb_none+=1


        if len(listes_data_depart[r])==nb_none:     # vide la liste remplie de non
            listes_data_depart[r].clear()
        nb_none=0

    for valeur in listes_data_depart :
        if len(valeur)!=0:
            listes_data_traiter.append(valeur)

    for titre in range(len(listes_data_traiter[0])):

        if listes_data_traiter[0][titre]=="Organisme d'accueil":
            for colonne in range(len(listes_data_traiter)):
                data_nom_entreprise.append(listes_data_traiter[colonne][titre])

        if listes_data_traiter[0][titre]=='Lieu':  # Si le nom de la colonne égal ADRENTR1 alors on parcourt ligne ADRENTR1 et on ajoute chaque adrsse d'entreprise dans la liste data_adresse
            for colonne in range(len(listes_data_traiter)):
                data_lieu.append(listes_data_traiter[colonne][titre])

        if listes_data_traiter[0][titre]=='Sujet':                          # Si le nom de la colonne égal LOCALITE alors on parcourt ligne LOCALITE et on ajoute chaque ville dans la liste data_ville
            for colonne in range(len(listes_data_traiter)):
                data_sujet.append(listes_data_traiter[colonne][titre])

        if listes_data_traiter[0][titre]=='Tuteur entreprise':                             # Si le nom de la colonne égal Pays alors on parcourt ligne Pays et on ajoute chaque pays dans la liste data_pays
            for colonne in range(len(listes_data_traiter)):
                data_tuteur.append(listes_data_traiter[colonne][titre])
            data_tuteur.pop(0)
        if listes_data_traiter[0][titre]=='Contact tel':
            for colonne in range(len(listes_data_traiter)):
                data_tel.append(listes_data_traiter[colonne][titre])

        if listes_data_traiter[0][titre]=='Mail':
            for colonne in range(len(listes_data_traiter)):
                data_mail.append(listes_data_traiter[colonne][titre])


    for i in range(len(data_tuteur)):
        print(data_tuteur[i][:1])
        if data_tuteur[i][:2] == 'Mr':
            print('lo')
            lettre_maj+=3
        while data_tuteur[i][lettre_maj].isupper()!=False:
            lettre_maj+=1
        
        
        data_nom_tuteur.append(data_tuteur[i][:int(lettre_maj)])
        data_prenom_tuteur.append(data_tuteur[i][int(lettre_maj)+1:])
        lettre_maj=0

    workbook.close()
    return data_nom_tuteur,data_prenom_tuteur

valeur_tmp=fichier_excel('Entreprises_stage_21_22.xlsx')




def contact(new_file):
    all_value=valeur_tmp
    print(all_value)
    

    

    


    if new_file[len(new_file)-4:] == 'xlsx':
        """
        Toute les colones à faire dans le fichiers .xlsx
        colone nom et prenom
        colonne entreprise
        colonne fonction
        colonne email
        colonne e-mail
        colonne téléphone
        """
        wb_out = openpyxl.Workbook()
        #accès à la première feuille du fichier
        ws1 = wb_out.active
        ws1.title = "Etudiants"

        # ajout entête des colonnes
        ws1.append(["Nom", "Prénom", "Note"])
        

        for row in donnees:
            ws1.append(row)

            #écriture du fichier
        wb_out.save(filename = 'out.xlsx')


    if new_file[len(new_file)-5:] == 'vcard':

        """
        faire un format txt
        NOM;
        PRENOM;
        ENTR;
        POSTE;
        EMAIL;
        TEL;

        

        for valeur in range(len(all_value)):
            #with open(new_file, 'w') as writefile:
                #writefile.write("BEGIN:VCARD\n")
            for data in range(len(valeur)):
                    #writefile.write(contenu_cvf[data]+valeur[data]+"\n")
                
        """          

def page(test):
    """Modifier os par de la manipulation txt"""

    filein = open(sys.argv[1], "r")
    fileout = open("html-table.html", "w")
    data = filein.readlines()

    table = "<table>\n"

    # Create the table's column headers
    header = data[0].split(",")
    table += "  <tr>\n"
    for column in header:
        table += "    <th>{0}</th>\n".format(column.strip())
    table += "  </tr>\n"

    # Create the table's row data
    for line in data[1:]:
        row = line.split(",")
        table += "  <tr>\n"
        for column in row:
            table += "    <td>{0}</td>\n".format(column.strip())
        table += "  </tr>\n"

    table += "</table>"

    fileout.writelines(table)
    fileout.close()
    filein.close()





