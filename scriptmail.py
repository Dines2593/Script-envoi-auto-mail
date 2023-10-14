from email import encoders
from email.mime.base import MIMEBase
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import re 

# Lecture du fichier Excel
excel_file_path = "./fichier_final.xlsx" 
df = pd.read_excel(excel_file_path)

# Adresse e-mail et mot de passe de l'expéditeur
sender_email = "ndinespro@gmail.com"
sender_password = "rbdolakklgmwtorx"

with open("./message.txt", "r", encoding='utf-8') as file: 
    message_content = file.read()


def is_valid_email(email):
    # Utilisation d'une expression régulière pour vérifier le format de l'e-mail
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

# Chemin vers ton CV (fichier PDF dans cet exemple)
cv_file_path = "./CV Nagulanathan Dines.pdf" 

# Fonction pour envoyer un e-mail avec pièce jointe
def send_email_with_attachment(to_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Ajout de la pièce jointe
    with open(attachment_path, "rb") as cv_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(cv_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename=CV.pdf')
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
        print(f"E-mail avec pièce jointe envoyé à {to_email}")
    except Exception as e:
        print(f"Erreur lors de l'envoi à {to_email}: {str(e)}")

# ... (boucle sur chaque ligne du fichier Excel reste inchangée)

# Modifier l'appel à la fonction pour envoyer l'e-mail
for index, row in df.iterrows():
    email = row['Courriel']
    if is_valid_email(email):
        send_email_with_attachment(email, "[CANDIDATURE Spontanée] - Alternance Informatique", message_content, cv_file_path)
    else:
        print("L'email est vide, ne peut pas envoyer l'e-mail.")
        
"""send_email_with_attachment("dines2593@gmail.com", "[CANDIDATURE] - Alternance Informatique", message_content, cv_file_path)"""