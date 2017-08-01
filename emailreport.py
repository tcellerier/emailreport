#!/usr/bin/python
# -*- coding: utf-8 -*-

################################
# Script de mailing automatisé #
################################
#
# Thomas Cellerier
# v1.0 - 2017/08


# Set default encoding to utf-8
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

# Standard libraries
import argparse
import os.path
import datetime
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from email.MIMEBase import MIMEBase
from email import encoders
import StringIO

# Libraries à installer (avec pip)
try:
    import numpy as np
    print "numpy version " + np.__version__
except ImportError: 
    print "\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n!!!!! Library numpy not installed !!!!!\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"
try:
    import pandas as pd  
    print "pandas version " + pd.__version__
except ImportError: 
    print "\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n!!!!! Library pandas not installed !!!!!\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"

try:
    import xlsxwriter 
    print "xlsxwriter version " + xlsxwriter.__version__
except ImportError: 
    print "\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n!!!!! Library xlsxwriter not installed !!!!!\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"
#import openpyxl # autre library Excel possible

try:
    import matplotlib
    import matplotlib.pyplot as plt # sudo apt-get install python-matplotlib
    print "matplotlib version " + matplotlib.__version__
except ImportError:
    print "\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n!!!!! Library matplotlib not installed !!!!!\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n"





#######################################################
##############     Class emailreport     ##############
############   Genère et envoi un email   #############
#######################################################

class emailreport(object):

    # emailTo, emailCC :  doit être un string avec les destinataires séparés par des virgules
    def __init__(self, emailSubject, emailFrom, emailTo, emailCC = ''):
        
        #  args.email : override du destinaire si option --email activée au lancement du script
        try: 
            args.email
        except NameError:
            pass
        else: # si existe
            if not args.email is None:
                emailTo = args.email
                emailCC = ''

        # MIME ultipart Related: The message consists of a root part (by default, the first) which reference other parts inline, which may in turn reference other parts. 
        # Multipart/mixed is used for sending files with different "Content-Type" headers inline (or as attachments). If sending pictures or other easily readable files, most mail clients will display them inline (unless otherwise specified with the "Content-disposition" header). Otherwise it will offer them as attachments. 
        emailmsg = MIMEMultipart('mixed')
        emailmsg['Subject'] = emailSubject
        emailmsg['From'] = emailFrom
        emailmsg['To'] = emailTo # doit être un string avec les destinataires séparés par des virgules
        emailmsg['CC'] = emailCC # doit être un string avec les destinataires séparés par des virgules

        self.emailmsg = emailmsg



    # Joint un fichier csv à self.emailmsg contenant les données de dataframe
    def attach_csv_generic(self, filename, dataframe):

        # Création du csv dans un buffer
        csvbuff = StringIO.StringIO()
        dataframe.to_csv(path_or_buf=csvbuff, sep=';', na_rep='', decimal=',', encoding='utf-8', index=False)

        # Attachement du csv à l'email
        EmailCsv = MIMEBase('application', 'octet-stream')
        EmailCsv.set_payload(csvbuff.getvalue())
        EmailCsv.add_header('Content-Disposition', 'attachment', filename=filename)
        self.emailmsg.attach(EmailCsv)
        csvbuff.close()


    # Joint un fichier Xlsx à self.emailmsg contenant les données de dataframe
    def attach_xlsx_generic(self, filename, dataframe):

        # Création de l'excel dans un buffer avec panda et xlsxwriter
        xlsxbuff = StringIO.StringIO()
        xlsxwriter = pd.ExcelWriter(xlsxbuff, engine='xlsxwriter')
        dataframe.to_excel(excel_writer=xlsxwriter, sheet_name='Sheet1', na_rep='', index=False)
        wb = xlsxwriter.book
        wb.set_properties({'title': 'Automatic Email report', 'author': 'Thomas Cellerier', 'company':  ''})

        # Attachement du xlsx à l'email
        xlsxwriter.save()
        EmailXlsx = MIMEBase('application', 'octet-stream')
        EmailXlsx.set_payload(xlsxbuff.getvalue())
        encoders.encode_base64(EmailXlsx)
        EmailXlsx.add_header('Content-Disposition', 'attachment', filename=filename)
        self.emailmsg.attach(EmailXlsx)
        xlsxbuff.close()


    # Méthode d'ajout du corps du message
    def add_body(self, emailbody):

        # On ajoute le corps du message à l'email
        self.emailmsg.attach(emailbody)


    #  Méthode d'envoi de l'email 
    #    emailErrorTo: destinaire de l'email d'exception en cas d'erreur
    def send_email(self, emailErrorTo = 'alice@gmail.com'):

        # envoi de l'email
        try:
            smtpObj = smtplib.SMTP('smtp-01.rbx')
            smtpObj.sendmail(self.emailmsg['From'], self.emailmsg['To'].split(',') + self.emailmsg['CC'].split(','), self.emailmsg.as_string())    
            print "\n--- Email sent successfully ---\n    Subject: %s\n    From: %s\n    To: %s" % (self.emailmsg['Subject'], self.emailmsg['From'], self.emailmsg['To'])
            if (self.emailmsg['CC'] != ''):
                print "    CC: %s" % self.emailmsg['CC']
            smtpObj.quit()

        # Si erreur lors de l'envoi, alerte par email et retourne fail
        except Exception as e:
            ErrorMsg =  "Subject: %s\nFrom: %s\nTo: %s\nCC: %s\n\nEmail not sent ! \nException: %s \n\nDatetime script run: %s \nScript: %s" % (self.emailmsg['Subject'], self.emailmsg['From'], self.emailmsg['To'],  self.emailmsg['CC'], e, datetimestart.isoformat(' '), args)
            print "\n%s \n\n" % ErrorMsg
            EMailError = MIMEText(ErrorMsg)
            EMailError['Subject'] = "Error - %s" % self.emailmsg['Subject']
            EMailError['From'] = self.emailmsg['From']
            EMailError['To'] = emailErrorTo
            smtpObj = smtplib.SMTP('smtp-01.rbx')
            smtpObj.sendmail(self.emailmsg['From'], emailErrorTo, EMailError.as_string())    
            smtpObj.quit()
            raise SystemExit(1) # On sort du script (pour éviter les spams si multiples envois). Return failure




#############################################
#######     Fonctions Transverses     #######
#############################################

# Retourne un tableau HTML contenant les données de dataframe au format styleCSS 
# Style CSS appliqué par défaut au tableau : dataframe
# DisplayIndex : display index (useful for pivot table)
# Formatters : format des colonnes. Ex: {'CLICKS_DW' : lambda x: "*%f*" % x if x<100 else str(x)}
def convert_df_to_html(df, DisplayIndex = False, formattersTable = None):

    pd.set_option('display.max_colwidth', -1) # sinon par defaut, on n'a que 50 caractères max par champ
    # Encode en string utf8
    # on supprime le border=1 du tableau, styleCSS par défaut : dataframe
    
    return df.to_html(formatters=formattersTable, header=True, index=DisplayIndex, na_rep=' ', bold_rows=False, escape=False).replace('border="1"','').replace('None','&nbsp;')


# Encapsuler du texte ou HTML dans un MIME multipart/related afin d'afficher les images inline sans les proposer en pièce jointe
#    format_body = html ou plain
def initialize_body_email(body, format_body):

    # Consider an image you want to display in the email body but don't want to offer to the recipient for download. 
    # You would need to use a related message to wrap the email body and attach the related message to the original mixed message.
    #   https://www.anomaly.net.au/blog/constructing-multipart-mime-messages-for-sending-emails-in-python/
    wrap_related_message = MIMEMultipart('related')
    wrap_related_message.attach(MIMEText(body, format_body, _charset='utf-8'))

    return wrap_related_message


# Fonction pour inclure une image du serveur dans l'email
#   in_html_body: 1 => image inclus dans le corps HTML avec le code <img src="cid:imgEmailName">  (avec les doubles quotes et pas single quote)
#                 0 => image en pièce jointe
def include_local_image(imgRelativePath, imgEmailName, emailmsg, in_html_body = 1):

    # Open image 
    with open(imgRelativePath, 'rb') as fp:
        msgImage = MIMEImage(fp.read())

        if in_html_body == 1: # Cas d'une image dans le corps HTML
            msgImage.add_header('Content-ID', '<' + imgEmailName + '>' ) # Define the image's ID as referenced above
        else: # sinon en pièce jointe
            msgImage.add_header('Content-Disposition', 'attachment', filename=imgEmailName)
        
        emailmsg.attach(msgImage)



# Fonction pour attacher les figures (bufferImg) au corps html (emailbody_html) avec un nom (imageName)
#   insertion dans code HTML :  <img src="cid:imageName"> (avec double quote obligatoire pour compatibilité)
def attach_figure(bufferImg, imageName, emailbody_html):
    msgFig = MIMEImage(bufferImg.getvalue())
    msgFig.add_header('Content-ID', '<' + imageName + '>' )
    emailbody_html.attach(msgFig)   
    bufferImg.close()




######################################################
##################  Class csstable  ##################
# generate email body with a nice CSS table template #
######################################################

class CSSTableBody(object):

    def __init__(self, title, tableWidth='800'):
        self.add_header(title, tableWidth)


    # generate HTML header 
    def add_header(self, title, tableWidth='800'):

        self.body_html = """
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
   "http://www.w3.org/TR/html4/loose.dtd">
<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>Email report</title>
    <style type="text/css">
      table.dataframe  {
        border-collapse: collapse;
        width: 100%%;
      }
      table.dataframe, table.dataframe th, table.dataframe td {
         border: 1px solid #000;
      }
      table.dataframe th, table.dataframe td {
         padding: 2px 3px; 
      }  
      table.dataframe thead, table.dataframe thead thead  { 
        text-align: center;
        background-color: #CCCCCC;
        font: 13px Arial, Helvetica, sans-serif;
      }
      table.dataframe td { 
        text-align: left;
        vertical-align: middle;
        background-color: white;
        font: 12px Arial, Helvetica, sans-serif;
      }  

      hr { 
        margin: 10px 0; 
        padding: 0; 
        height: 1px; 
        background-color: #fff; 
        border: none; 
        border-top: 1px solid #000;
      }
    </style>
  </head>
  <body>
    <!-- container -->
    <!-- header -->
    <table width="%s" align="center" bgcolor="#ececeb" border="0" cellpadding="6" cellspacing="0">
        <tr>
            <td>
                <table width="100%%" bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td width="100%%" style="padding: 10px 0;">
                            <!-- header -->
                            <table width="100%%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td width="215" align="left" valign="middle">
                                        <font face="Arial, Helvetica, sans-serif">Logo</font>
                                    </td>
                                    <td valign="middle">
                                        <font face="Arial, Helvetica, sans-serif" size="4">&nbsp;&nbsp;&nbsp;%s</font>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td width="100%%" height="20" bgcolor="#ececeb" valign="top">
                            <span style="display:block; background-color:#2ea5d3;"><img src="cid:emailborder" width="%s" height="14" border="0" align="top" alt="" style="display:block;"></span>
                        </td>
                    </tr>
<!-- /header -->
<!-- content -->
""" % (tableWidth, title, tableWidth)



    # generate HTML table  
    # DisplayIndex : display index (useful for pivot table)
    # Formatters : format des colonnes. Ex: {'CLICKS_DW' : lambda x: "*%f*" % x if x<100 else str(x)}
    def add_table(self, df, titleKPI, DescriptionKPI = '', DisplayIndex = False, formattersTable = None):

        table_html = """
                    <tr><td bgcolor="#FFFFFF" style="font-family: Arial, Helvetica, sans-serif; padding: 10px;">
                      <table width="100%" border="0" cellpadding="0" cellspacing="0" style="padding-bottom: 10px;"><tr><td style="padding: 7px; font-family: Arial, Helvetica, sans-serif; font-size: 16px; color: #333333; font-weight: bold; background-color: #ececeb;">{0}</td></tr></table>
                      <p style="margin-top: 5px; font-size: 14px;">{1}</p>
        """.format(titleKPI, DescriptionKPI)
        
        table_html += convert_df_to_html(df, DisplayIndex=DisplayIndex, formattersTable=formattersTable)
        table_html += "<hr>\n        </td></tr> \n"

        self.body_html += table_html


    # generate HTML footer  
    def add_footer(self):
        self.body_html += """
    <!-- footer -->
                    <tr>
                        <td bgcolor="#ececeb" style="padding: 10px 0;">
                            <font face="Arial, Helvetica, sans-serif" size="2" color="#999999">
                                <a style="font-weight: bold; color:#999999; text-decoration:none;" href="mailto:alice@gmail.com">Alice</a><br/>
                            </font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <!-- /footer -->
    <!-- /container -->
    </body>
    </html>"""


    # Création de l'encapsulation de l'HTML et de l'ajout des images dans le corps du message
    def generate_email_body(self):

        email_body_Html = initialize_body_email(self.body_html, 'html')
        include_local_image('img/emailBorder.png', 'emailborder', email_body_Html)

        return email_body_Html







#############################################################################
########                     Exécution du script                     ########
#############################################################################

if __name__ == "__main__":

    ##############
    #   Header   #
    ############## 

    datetimestart = datetime.datetime.now()
    print """# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
        _____                 _ _      ____                       _
       | ____|_ __ ___   __ _(_) |    |  _ \ ___ _ __   ___  _ __| |_
       |  _| | '_ ` _ \ / _` | | |    | |_) / _ \ '_ \ / _ \| '__| __|
       | |___| | | | | | (_| | | |    |  _ <  __/ |_) | (_) | |  | |_
       |_____|_| |_| |_|\__,_|_|_|    |_| \_\___| .__/ \___/|_|   \__|
                                                |_|
        Version 1.0 - 08/2017

"""

    ####################################
    # Lecture des paramètres en entrée #
    ####################################

    parser = argparse.ArgumentParser()
    parser.add_argument("script_file", help="Specify the script file to execute (inside ./scripts/ directory, without .py)")
    parser.add_argument("-e", "--email", dest="email", metavar="<email>", help="Override email recipient")
    parser.add_argument("-p1", "--param1", dest="param1", metavar="<param1>", help="Parameter #1 to send to script")
    parser.add_argument("-p2", "--param2", dest="param2", metavar="<param2>", help="Parameter #2 to send to script")
    parser.add_argument("-p3", "--param3", dest="param3", metavar="<param3>", help="Parameter #3 to send to script")
    args = parser.parse_args()


    #######################
    # Exécution du script #
    #######################

    # On vérifie que le script demandé existe bien puis on l'exécute
    if os.path.isfile("./scripts/" + args.script_file + ".py"):
        execfile("./scripts/" + args.script_file + ".py")
    else:
        print "!!!!    Error: File ./scripts/%s.py does not exist    !!!!\n" % args.script_file 
        raise SystemExit(1) # Return failure


    ##############
    #   Footer   #
    ############## 

    # Temps d'exécution
    datetimediff = datetime.datetime.now() - datetimestart

    print """\nExecution time: %s \n
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
""" % datetimediff
