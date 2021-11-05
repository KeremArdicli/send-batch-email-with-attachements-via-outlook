import csv
from time import sleep
import win32com.client as client

with open('mailbody.html', 'r', encoding='utf-8') as f:     # prepare an html file named mailbody.htm and but it in the same directory.
    html_string = f.read() 
    
template = html_string 
# template = "{} , Mail body here."  ---------------- to use plain text activate this line and deactivate the line above.

with open('people.csv', 'r', newline='') as f:
     reader = csv.reader(f)
     distro = [row for row in reader]


chunks = [distro[x:x+4] for x in range (0, len(distro), 4)]   #change the numbers in this line according to the number of rows in your csv file. for eg: if you have 20 people, change 4s to 20.
outlook = client.Dispatch('Outlook.Application')

for chunk in chunks:
    for name, email, attach in chunk:
        message = outlook.CreateItem(0)
        message.To = email
        message.Subject = "Mail Subject Here"
        message.Body = template.format(name) # for html templates, change this line to "message.HTMLBody = template"
        attachement = attach
        message.Attachments.Add(attachement)
        message.Send()
    sleep(60)
