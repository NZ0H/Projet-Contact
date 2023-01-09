def fusion(data):
    """
     Création : Enzo, 06/01/2023
     Modification : 07/01/2023



     
    """
    #liste avec toute les données assemblées
    data_nom_entreprise=[]
    data_ville=[]
    data_code_postal=[]
    data_sujet=[]
    data_civilite_tuteur=[]
    data_nom_tuteur=[]
    data_prenom_tuteur=[]
    data_tel=[]
    data_mail=[]
    #liste permetant de trier les doublons
    entreprise=[]
    ville=[]
    code_postal=[]
    sujet=[]
    civilite_tuteur=[]
    nom_tuteur=[]
    prenom_tuteur=[]
    tel=[]
    mail=[]

    #trie toute les données recuent dans chaques listes
    for fichier in range(len(data)):
        for liste in range(len(data[fichier])):
            if liste == 0 :
                for valeur in data[fichier][liste]:
                    data_nom_entreprise.append(valeur)
            if liste == 1 :
                for valeur in data[fichier][liste]:
                    data_ville.append(valeur)
            if liste == 2 :
                for valeur in data[fichier][liste]:
                    data_code_postal.append(valeur)
            if liste == 3 :
                for valeur in data[fichier][liste]:
                    data_sujet.append(valeur)
            if liste == 4 :
                for valeur in data[fichier][liste]:
                    data_civilite_tuteur.append(valeur)
            if liste == 5 :
                for valeur in data[fichier][liste]:
                    data_nom_tuteur.append(valeur)
            if liste == 6 :
                for valeur in data[fichier][liste]:
                    data_prenom_tuteur.append(valeur)
            if liste == 7 :
                for valeur in data[fichier][liste]:
                    data_tel.append(valeur)
            if liste == 8 :
                for valeur in data[fichier][liste]:
                    data_mail.append(valeur)

    #ajoute que si le contenue n'est pas present dans la liste
    for num in range(len(data_nom_entreprise)):
        if (data_nom_entreprise[num] not in entreprise) and (data_nom_tuteur[num] not in nom_tuteur) and (data_tel[num] not in tel):
            entreprise.append(data_nom_entreprise[num])
            ville.append(data_ville[num])
            code_postal.append(data_code_postal[num])
            sujet.append(data_sujet[num])
            civilite_tuteur.append(data_civilite_tuteur[num])
            nom_tuteur.append(data_nom_tuteur[num])
            prenom_tuteur.append(data_prenom_tuteur[num])
            tel.append(data_tel[num])
            mail.append(data_mail[num])

    return entreprise,ville,code_postal,sujet,civilite_tuteur,nom_tuteur,prenom_tuteur,tel,mail

