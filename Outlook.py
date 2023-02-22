# PASSO 1
import win32com.client as win32
outlook = win32.Dispatch('outlook.applicantion')

# PASSO 2
mail = outlook.CreateItem(0)
mail.To = 'arthursgrasso@gmail.com'
mail.Subject = 'Email vindo do Outlook'
mail.Body = 'Texto do E-mail'

attachement = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs'
mail.Attachments.Add(attachement)

mail.Send()