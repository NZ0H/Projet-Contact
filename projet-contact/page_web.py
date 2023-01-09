import os 
from fusion import fusion

def page_web(args,valeur_tmp_fct2):  
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
