import openpyxl

def fusion(data):
    data_nom_entreprise=[]
    data_ville=[]
    data_code_postal=[]
    data_sujet=[]
    data_civilite_tuteur=[]
    data_nom_tuteur=[]
    data_prenom_tuteur=[]
    data_tel=[]
    data_mail=[]
    
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
    print
    return data_nom_entreprise,data_ville,data_code_postal,data_sujet,data_civilite_tuteur,data_nom_tuteur,data_prenom_tuteur,data_tel,data_mail

