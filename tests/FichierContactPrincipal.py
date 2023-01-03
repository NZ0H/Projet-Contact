import openpyxl
import argparse

#Fichier pour tester la suppression des lignes None
"""
def argument():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file',type=str,nargs='+')
    parser.add_argument('--sortie',type=str)
    parser.add_argument('--nom_page',type=str)
    args=parser.parse_args()
    return args

"""





def fichier_excel(args):
    workbook = openpyxl.load_workbook('../data/Stages/'+args, data_only = True)
    titres_onglets = workbook.sheetnames
    onglet = workbook[titres_onglets[0]]

    nb_none=0

    listes_data_depart = [] #stocke toute les lignes
    listes_data_traiter = []

    """data recupérer directement dans le excel, sans avoir à faire de trie"""

    data_nom_entreprise=[]
    data_lieu=[]
    data_sujet=[]
    data_tuteur=[]
    data_tel=[]
    data_mail=[]

    """ data ayant était traité"""
    data_code_postal=[]

    data_prenom_tuteur=[]
    data_nom_tuteur=[]
    data_civilite_tuteur=[]
    data_code_postal=[]
    data_ville=[]

    nom_prenom_tmp=[]
    nom_tmp=[]
    prenom_tmp=[]
    valeur_tmp=[]
    valeur=[]


    for row in onglet.values: #ajoute les lignes dans la liste lignes

        listes_data_depart.append(list(row))

    for r in range(len(listes_data_depart)):        #parcourt les listes présente dans lignes
        for non in listes_data_depart[r]:       #parcourt liste de liste pour avoir les valeurs individuelles
            if non == None:
                nb_none+=1
        if len(listes_data_depart[r])==nb_none:
            listes_data_depart[r].clear()
        nb_none=0

    for valeur in listes_data_depart :
        if len(valeur)!=0:
            listes_data_traiter.append(valeur)

    for titre in range(len(listes_data_traiter[0])):

        if listes_data_traiter[0][titre]=="Organisme d'accueil":
            for colonne in range(len(listes_data_traiter)):
                data_nom_entreprise.append(listes_data_traiter[colonne][titre])
            data_nom_entreprise.remove("Organisme d'accueil")
        if listes_data_traiter[0][titre]=='Lieu':
            for colonne in range(len(listes_data_traiter)):
                data_lieu.append(listes_data_traiter[colonne][titre])
            data_lieu.remove('Lieu')
        if listes_data_traiter[0][titre]=='Sujet':
            for colonne in range(len(listes_data_traiter)):
                data_sujet.append(listes_data_traiter[colonne][titre])
            data_sujet.remove('Sujet')
        if listes_data_traiter[0][titre]=='Tuteur entreprise':
            for colonne in range(len(listes_data_traiter)):
                data_tuteur.append(listes_data_traiter[colonne][titre])
            data_tuteur.remove('Tuteur entreprise')
        if listes_data_traiter[0][titre]=='Contact tel':
            for colonne in range(len(listes_data_traiter)):
                data_tel.append(listes_data_traiter[colonne][titre])
            data_tel.remove('Contact tel')
        if listes_data_traiter[0][titre]=='Mail':
            for colonne in range(len(listes_data_traiter)):
                data_mail.append(listes_data_traiter[colonne][titre])
            data_mail.remove('Mail')
            
     
    
    for i in range(len(data_tuteur)):
        if data_tuteur[i]==None:
            nom_prenom_tmp.append(data_tuteur[i])
        else:
            nom_prenom_tmp.append(data_tuteur[i].split())

    print(nom_prenom_tmp)

    for info in range(len(nom_prenom_tmp)):
        
            data_civilite_tuteur.append(nom_prenom_tmp[info][0])
            del nom_prenom_tmp[info][0]
            for n_p in range(len(nom_prenom_tmp[info])):

                if nom_prenom_tmp[info][n_p].isupper() == True :
                    nom_tmp.append(nom_prenom_tmp[info][n_p])
                else :
                    prenom_tmp.append(nom_prenom_tmp[info][n_p])

            data_nom_tuteur.append(" ".join(nom_tmp))
            nom_tmp.clear()
            data_prenom_tuteur.append(" ".join(prenom_tmp))
            prenom_tmp.clear()
    

    """separation code postal et ville"""
    
    for i in range(len(data_lieu)):
        data_code_postal.append('Champs vide')
        valeur_tmp.append(data_lieu[i].split())

        for v in range(len(valeur_tmp[i])):

            """cree liste avec valeur vide"""

            if valeur_tmp[i][v].isdecimal() == True :
                data_code_postal[i]=valeur_tmp[i][v]

            if valeur_tmp[i][v].isalpha() == True :
                valeur.append(valeur_tmp[i][v])

        data_ville.append(" ".join(valeur))
        valeur.clear()
    
    workbook.close()
    print(len(data_sujet),len(data_civilite_tuteur))
    return data_nom_entreprise,data_ville,data_code_postal,data_sujet,data_civilite_tuteur,data_nom_tuteur,data_prenom_tuteur,data_tel,data_mail
print(fichier_excel('Entreprises_stage_2021.xlsx'))
valeur_tmp=[]

"""

def contact(args_sortie):
    all_value=valeur_tmp
    data_nom_entreprise=[]
    data_ville=[]
    data_code_postal=[]
    data_sujet=[]
    data_civilite_tuteur=[]
    data_nom_tuteur=[]
    data_prenom_tuteur=[]
    data_tel=[]
    data_mail=[]
    print(all_value)
   
    for fichier in range(len(all_value)):
        for liste in range(len(all_value[fichier])):
            print(len(all_value[fichier]))
            if liste == 0 :
                print('ok')
                for valeur in all_value[liste][0]:
                    data_nom_entreprise.append(valeur)
                print(data_nom_entreprise)

            if liste == 1 :
                print('ok')
                for valeur in all_value[liste][0]:
                    print(all_value[liste][0])
                    data_ville.append(valeur)

            if liste == 2 :
                print('ok')
                for valeur in all_value[liste][0]:
                    print(all_value[liste][0])
                    data_code_postal.append(valeur)

            if liste == 3 :
                print('ok')
                for valeur in all_value[liste][0]:
                    print(all_value[liste][0])
                    data_sujet.append(valeur)

            if liste == 4 :
                print('ok')
                for valeur in all_value[liste][0]:
                    data_civilite_tuteur.append(valeur)

            if liste == 5 :
                print('ok')
                for valeur in all_value[liste][0]:
                    data_nom_tuteur.append(valeur)

            if liste == 6 :
                for valeur in all_value[liste][0]:
                    data_prenom_tuteur.append(valeur)

            if liste == 7 :
                for valeur in all_value[liste][0]:
                    data_tel.append(valeur)
            if liste == 8 :
                for valeur in all_value[liste][0]:
                    data_mail.append(valeur)
    print(data_mail)
    if args_sortie[len(args_sortie)-4:] == 'xlsx':
        print(data_code_postal)
    
      
      
        wb_out = openpyxl.Workbook()
        #accès à la première feuille du fichier
        ws1 = wb_out.active
        ws1.title = "Etudiants"

        # ajout entête des colonnes
        ws1.append(["Nom Entreprise","Ville","Code Postal","Sujet","Civilité","Nom","Prénom","Téléphone","Email"])


        for row in len(data_nom_entreprise):

            ws1.append(data_nom_entreprise[row])
            ws1.append(data_ville[row])
            ws1.append(data_sujet[row])
            ws1.append(data_civilite_tuteur[row])
            ws1.append(data_nom_tuteur[row])
            ws1.append(data_prenom_tuteur[row])
            ws1.append(data_tel[row])   
            ws1.append(data_mail[row]) 
        
            

        
        wb_out.save(filename = new_file)
  """


#arguments=argument()

#for arg in arguments.file :
    #valeur_tmp.append(fichier_excel(arg))


#print(contact(arguments.sortie))

