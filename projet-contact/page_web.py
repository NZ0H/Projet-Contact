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

    f.write("<tr>")

    for i in range(len(data[0])):
        for cell in [data[0][i], data[1][i],data[2][i],data[4][i], data[5][i], data[6][i], data[3][i],data[8][i], data[7][i]]:
            f.write(f"<td>{cell}</td>")
        f.write("</tr>")

    f.write("</table>\n</body>\n</html>")
    f.close()
