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
# # #                        Template TEXT                        # # #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
"""


#####################
# Paramétrage email #
#####################

emailFrom = 'alice@gmail.com'
emailTo = 'bob@gmail.com' # doit être un string avec les destinataires séparés par des virgules, et different de From
emailCC = ''  # optionnel (string avec les destinataires séparés par des virgules)
emailErrorTo = 'error@gmail.com'  # doit etre different de From

emailSubject = "Test txt"

# Initialisation email
emailreport_template = emailreport(emailSubject, emailFrom, emailTo, emailCC)

# Création du corps de l'email
body_txt = "Test\nFormat plain text"
emailbody_txt = initialize_body_email(body_txt, 'plain')
emailreport_template.add_body(emailbody_txt)

# pièce jointe image (qu'on attache au message root)
include_local_image('img/test.png', 'test.png', emailreport_template.emailmsg, in_html_body=0) 

# Envoi de l'email
emailreport_template.send_email(emailErrorTo)
