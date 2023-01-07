import os

data_prenom_tuteur = "Anthony"
data_nom_tuteur = "Chapus"
data_mail = "mail@gmail.com"
data_tel = "0606060606"
data_nom_entreprise = "Orange"
data_sujet = "Chef de projet"


os.mkdir("dossier2")
f = open("dossier2/data.vcf","w")
f.write("BEGIN:VCARD\nVERSION:3.0\n")
f.write("FN:{} {}\n".format(data_prenom_tuteur,data_nom_tuteur))
f.write("N:{};{};;;\n".format(data_nom_tuteur,data_prenom_tuteur))
f.write("item1.EMAIL;TYPE=INTERNET:{}\n".format(data_mail))
f.write("item1.X-ABLabel:\nitem2.TEL:{}\n".format(data_tel))
f.write("item2.X-ABLabel:\nitem3.ORG:{}\n".format(data_nom_entreprise))
f.write("item3.X-ABLabel:\nNOTE:{}\n".format(data_sujet))
f.write("CATEGORIES:myContacts\nEND:VCARD")