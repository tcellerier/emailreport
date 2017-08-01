#!/usr/bin/python
# -*- coding: utf-8 -*-


###################################
# Fonctions standards disponibles #
###################################

# Retourner un tableau HTML contenant les données de dataframe au format styleCSS :
#   convert_df_to_html(df, DisplayIndex = False, formattersTable = None)
# Encapsuler la partie HTML ou texte dans un MIME multipart/related afin d'afficher les images inline sans les proposer en pièce jointe :
#   initialize_body_email(body, format_body)
# Fonction pour inclure une image dans l'email :
#   include_local_image(imgRelativePath, imgEmailName, emailmsg, in_html_body = 1)

######### Dans la class emailreport #########
# Joindre un CSV :
#   emailreport.attach_csv_generic("fichier.csv", df)
# Joindre un xlsx :
#   emailreport.attach_xlsx_generic("toto1.xlsx", df)


import MySQLdb

print """# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# # #                    Template CSS Table                       # # #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
"""


#####################
# Paramétrage email #
#####################

emailFrom = 'alice@gmail.com'
emailTo = 'bob@gmail.com' # doit être un string avec les destinataires séparés par des virgules, et different de From
emailCC = '' # optionnel (string avec les destinataires séparés par des virgules)
emailErrorTo = 'error@gmail.com'  # doit etre different de From

emailSubject = "Test CSS Table"

# Initialisation email
emailreport_csstable = emailreport(emailSubject, emailFrom, emailTo, emailCC)


# generation du dataframe
KPI1_title =  "Title KPI 1"
KPI1_description =  "Description KPI 1" # Optionnel
conn = MySQLdb.connect('server', 'user', 'password', 'db')
KPI1_query = "SELECT DATE, CLIENT_NAME, TAILLE from CLIENTS where DATE = CURRENT_DATE - INTERVAL 1 DAY limit 10"
KPI1_df = pd.read_sql_query(KPI1_query, conn)

# generation du tableau CSS
bodyhtml_csstable = CSSTableBody(emailSubject, tableWidth='800')
bodyhtml_csstable.add_table(KPI1_df, KPI1_title, DescriptionKPI=KPI1_description, DisplayIndex=False)
bodyhtml_csstable.add_footer()

# Création du corps de l'email
emailbody = bodyhtml_csstable.generate_email_body()
emailreport_csstable.add_body(emailbody)

# Envoi de l'email
emailreport_csstable.send_email(emailErrorTo)
