#1° Bibilioteca: pywin32
import win32com.client as win32

#2° Criar a interação com o outlook.
outlook = win32.Dispatch('outlook.application')

#3° Criar um Email
email = outlook.CreateItem(0)

#4° Configurar as informações do seu e-mail.
email.To = "........@gmail.com"
email.Subject = "Assunto do que você que enviar, pra determinado E-mail."
email.HTMLBody = """
<p>Ex: Olá Kayke, aqui o código em Python!</p>

<p>Dano início a tecnologia Python para trabalhar, com 
ciência de dados.</p>

<p>Abs,</p>
<p>Kayke-Ti.</p>
"""
#5° Enviar o Email.
email.Send()
print("Email enviado com success").