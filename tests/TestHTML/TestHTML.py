stucture = ["Entreprise","Civilite","Nom","Prenom","Post","Email","Telephone"]


def genere_page_web(nom_fichier, titre_page, style, structure):
    # Création de la structure de la page web
    meta = "'utf-8'"
    f = open(nom_fichier,"w")
    f.write("<!DOCTYPE html>\n")
    f.write("<head> \n<meta charset={}>\n".format(meta))
    f.write("<link rel='stylesheet' href='{}'>".format(style))
    f.write("\n<title> {} </title> \n</head>".format(titre_page))
    f.write("\n<body>\n")

    # Création du tableau
    f.write("<table>\n<tr>")
    for i in structure:
        f.write("<th>{}</th>".format(i))
    f.write("</tr>")



   
   
    f.close()

genere_page_web("pageweb.html","test","style.css",stucture)