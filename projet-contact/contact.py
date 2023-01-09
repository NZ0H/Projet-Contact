import openpyxl
from fusion import fusion

def contact(args_sortie,valeur_tmp_fct2):
    """
     Création : 
     Modification




    """
    all_value=valeur_tmp_fct2
    colonne=fusion(all_value)
    
    if args_sortie[len(args_sortie)-4:] == 'xlsx':
        wb_out = openpyxl.Workbook()
        ws1 = wb_out.active
        ws1.title = "Etudiants"
        ws1.append(["Nom Entreprise","Ville","Code Postal","Sujet","Civilité","Nom","Prénom","Téléphone","Email"]) 
        for entreprise,ville,code_postal,sujet,civilite,nom,prenom,tel,email in zip(colonne[0],colonne[1],colonne[2],colonne[3],colonne[4],colonne[5],colonne[6],colonne[7],colonne[8]):
            ws1.append([entreprise,ville,code_postal,sujet,civilite,nom,prenom,tel,email])
        wb_out.save('../VCARD/'+args_sortie)

    if args_sortie[len(args_sortie)-5:] == 'vcard':
        for num_c in range(len(colonne[5])):
            f= open("../VCARD/vcard_{}_{}.vcf".format(colonne[5][num_c],colonne[6][num_c]),"w")
            f.write("BEGIN:VCARD\nVERSION:3.0\n")
            f.write("FN:{} {}\n".format(colonne[6][num_c],colonne[5][num_c]))
            f.write("N:{};{};;;\n".format(colonne[5][num_c],colonne[6][num_c]))
            f.write("item1.EMAIL;TYPE=INTERNET:{}\n".format(colonne[8][num_c]))
            f.write("item1.X-ABLabel:\nitem2.TEL:{}\n".format(colonne[7][num_c]))
            f.write("item2.X-ABLabel:\nitem3.ORG:{}\n".format(colonne[0][num_c]))
            f.write("item3.X-ABLabel:\nNOTE:{}\n".format(colonne[3][num_c]))
            f.write("CATEGORIES:myContacts\nEND:VCARD")