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

print """# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# # #                        Template HTML                        # # #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
"""


#####################
# Paramétrage email #
#####################

emailFrom = 'alice@gmail.com'
emailTo = 'bob@gmail.com' # doit être un string avec les destinataires séparés par des virgules, et different de From
emailCC = ''  # optionnel (string avec les destinataires séparés par des virgules)
emailErrorTo = 'alice@gmail.com'  # doit etre different de From


emailSubject = "Test html"

# Initialisation email
emailreport_template = emailreport(emailSubject, emailFrom, emailTo, emailCC)


bodyHtml = """
 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
   "http://www.w3.org/TR/html4/loose.dtd">
<html>
   <head>
      <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
      <title>Email Report</title>
   </head>
   <body>Test <font color="red">Email Report</font><br>Format <b>HTML</b><br><br><img src="cid:test.png"> </body></html>"""


# Création du corps de l'email
emailbody_html = initialize_body_email(bodyHtml, 'html')
emailreport_template.add_body(emailbody_html)

# pièce jointe image (qu'on attache au corps html)
include_local_image('img/test.png', 'test.png', emailbody_html, in_html_body=1) 

# Envoi de l'email
emailreport_template.send_email(emailErrorTo)
