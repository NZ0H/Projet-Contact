import openpyxl




def fichier_excel(args):
    workbook = openpyxl.load_workbook('../data/'+ args, data_only = True)
    titres_onglets = workbook.sheetnames
    onglet = workbook[titres_onglets[0]]
    nb_none=0
    listes_data_depart = [] 
    listes_data_traiter = []

    """data recupérer directement dans le excel, sans les trier (apres avoir passer l'étape des suppression de ligne de None"""
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


    for row in onglet.values: 
        listes_data_depart.append(list(row))

    for ligne in range(len(listes_data_depart)):   
        
        for non in listes_data_depart[ligne]:
            if non == None:
                nb_none+=1
        if len(listes_data_depart[ligne])==nb_none:
            listes_data_depart[ligne].clear()
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

    workbook.close()
    
    for i in range(len(data_tuteur)):
        if data_tuteur[i]==None:
            nom_prenom_tmp.append(data_tuteur[i])
        else:
            nom_prenom_tmp.append(data_tuteur[i].split())

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
                
    for i in range(len(data_nom_entreprise)):  
        if  len(data_civilite_tuteur)==0 or len(data_code_postal)==0 or len(data_mail)==0 or len(data_sujet)==0 or len(data_ville)==0 or len(data_tel)==0:
            if len(data_civilite_tuteur)==0 :
                for rep in range(len(data_nom_entreprise)):
                    data_civilite_tuteur.append('Non_renseigné')
            if len(data_code_postal)==0 :
                for rep in range(len(data_nom_entreprise)):
                    data_code_postal.append('Non_renseigné')
            if len(data_mail)==0 :
                for rep in range(len(data_nom_entreprise)):
                    data_mail.append('Non_renseigné')
            if len(data_sujet)==0 :
                for rep in range(len(data_nom_entreprise)):
                    data_sujet.append('Non_renseigné')
            if len(data_ville)==0 :
                for rep in range(len(data_nom_entreprise)):
                    data_ville.append('Non_renseigné')
            if len(data_tel)==0 :
                for rep in range(len(data_nom_entreprise)):
                    data_tel.append('Non_renseigné')

    return data_nom_entreprise,data_ville,data_code_postal,data_sujet,data_civilite_tuteur,data_nom_tuteur,data_prenom_tuteur,data_tel,data_mail