import openpyxl


def fichier_excel(args): 
    """
    Cette fonction permet de trier les fichiers excel sélectionné.

    :param args: Fichier excel.
    :type args: str
    :returns: Renvoie toutes les données triées.
    :rtype: str
    :raises: TypeError
    :example:

    (["Orange","SFR"],[["Paris","Poitiers"]],[["75000","86000"]],[["Fibre","Cable"]],[["Mr","Mr"]],[["Bop","Tope"]],[["Tom","Lucas"]],[["0607080910","0605040302"]],[["Bop.tom@gmail.com","Tope.lucas@gmail.com"]])

    """

    workbook = openpyxl.load_workbook('../data/'+ args, data_only = True)
    titres_onglets = workbook.sheetnames
    onglet = workbook[titres_onglets[0]]
    nb_none=0
    listes_data_depart = [] 
    listes_data_traiter = []

    #data recupérer directement dans le excel, sans les trier (apres avoir passer l'étape des suppression de ligne de None)
    data_nom_entreprise=[]
    data_lieu=[]
    data_sujet=[]
    data_tuteur=[]
    data_tel=[]
    data_mail=[]

    #data ayant était traité
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

    #ajouts des valeurs du fichier excel dans la liste 'listes_data_depart'
    for row in onglet.values: 
        listes_data_depart.append(list(row))

    #suppression du contenue de la liste remplie de None
    for ligne in range(len(listes_data_depart)):   
        for non in listes_data_depart[ligne]:
            if non == None:
                nb_none+=1
        if len(listes_data_depart[ligne])==nb_none:
            listes_data_depart[ligne].clear()
        nb_none=0

    #si longueur liste différent de 0 alors on ajoute à la listes_data_traiter
    for valeur in listes_data_depart :
        if len(valeur)!=0:
            listes_data_traiter.append(valeur)

    #Traitement des données celon les titres des colonnes, si le titre fait partie des conditions il est ajouté sinon il n'est pas prit en compte
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
    
    #traitement de data_tuteur, si None on ajoute None sinon on sépare le champ de donnée dans une liste
    for d_t in range(len(data_tuteur)):
        if data_tuteur[d_t]==None:
            nom_prenom_tmp.append(data_tuteur[d_t])
        else:
            nom_prenom_tmp.append(data_tuteur[d_t].split())
    #séparation de nomn,prenom,civilite
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
    #trie la ville est le code postal 
    for d_v in range(len(data_lieu)):
        data_code_postal.append('Champs vide')
        valeur_tmp.append(data_lieu[d_v].split())
        for v in range(len(valeur_tmp[d_v])):
            if valeur_tmp[d_v][v].isdecimal() == True :
                data_code_postal[d_v]=valeur_tmp[d_v][v]
            if valeur_tmp[d_v][v].isalpha() == True :
                valeur.append(valeur_tmp[d_v][v])
        data_ville.append(" ".join(valeur))
        valeur.clear()
        
    #si un des champs est vide on ajoute len(data_nom_entreprise) fois 'Nom_renseigné' , len(data_nom_entreprise) car l'objectif de ce (->)
    #(suite) programme est pour les stages en entreprise donc le champ est forcément remplie (meme si chaque champ de data fait la meme longueur)       
    for compteur in range(len(data_nom_entreprise)):  
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