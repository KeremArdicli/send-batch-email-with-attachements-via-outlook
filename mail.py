import csv
from time import sleep
import win32com.client as client

template = "{} , Mail body here." # you can use html templates as well

with open('people.csv', 'r', newline='') as f:
     reader = csv.reader(f)
     distro = [row for row in reader]


chunks = [distro[x:x+4] for x in range (0, len(distro), 4)]
outlook = client.Dispatch('Outlook.Application')

for chunk in chunks:
    for name, email, attach in chunk:
        message = outlook.CreateItem(0)
        message.To = email
        message.Subject = "Mail Subject Here"
        message.Body = template.format(name)
        attachement = attach
        message.Attachments.Add(attachement)
        message.Send()
    sleep(60)