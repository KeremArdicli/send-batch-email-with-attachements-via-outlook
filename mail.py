import csv
from time import sleep
import win32com.client as client

with open('mailbody.html', 'r', encoding='utf-8') as f:
    html_string = f.read()
    
template = html_string


with open('people2.csv', 'r', newline='') as f:
     reader = csv.reader(f)
     distro = [row for row in reader]

f= open("report.txt","w+")

chunks = [distro[x:x+4] for x in range (0, len(distro), 4)]  # 2 people per chunk
outlook = client.Dispatch('Outlook.Application')

i = 1

for chunk in chunks:
    for name, email, attach in chunk:
        try:
            sign = "someimage.png"  #url of your mail signature 
            message = outlook.CreateItem(0)
            message.To = email
            message.Cc  = 'example@example.com'
            message.Subject = "HTML dosyadan çekme denemesi"
            message.HTMLBody = template.format(name=name, sign=sign)  # change line to message.Body if you like to send plain text
            attachement = attach
            message.Attachments.Add(attachement)
            message.Send()
            i += 1
        except:
            print("Mail could not sent to: ", name)
            f.write("Mail could not sent to: " + name + "\n")
    sleep(3)
    
f.write(str(i) + " e-mails has been sent..")
f.close()
