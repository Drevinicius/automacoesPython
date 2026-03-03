import smtplib
from email.message import EmailMessage
import ssl

# Configurações do e-mail
smtp_server = "smtp.gmail.com" # Server que cuidará do envio
port = 587 # Porta de saída padrão
sender_email = "anvnc01@gmail.com" # Remetente
receiver_email = "andrevinicius.magno@gmail.com" # Destinatário
password = "erka inwv tzsu gout"  # Use uma senha de app se necessário

codigo = 'EU SOU UM DEV FODA'

# Corpo do e-mail
msg = EmailMessage()
msg["Subject"] = "Seu código de verificação"
msg["From"] = sender_email
msg["To"] = receiver_email
msg.set_content(f"Olá!\n\nSeu código de verificação é: {codigo}\n\nNão responda este e-mail.", charset="utf-8")

# Enviar o e-mail
try:
    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=context)
        server.login(sender_email, password)
        server.send_message(msg)
    print("Email enviado com sucesso!")
except Exception as e:
    print(f"Erro ao enviar o email: {e}")# Configurações do e-mail
