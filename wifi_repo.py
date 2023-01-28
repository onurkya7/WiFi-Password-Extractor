import subprocess
from openpyxl import *

repo = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode('utf-8', errors="backslashreplace").split('\n')
users = [i.split(":")[1][1:-1] for i in repo if "All User Profile" in i]
list = []

#get password
for i in users:
    try:
        results = subprocess.check_output(['netsh', 'wlan', 'show', 'profile', i, 'key=clear']).decode('utf-8', errors="backslashreplace").split('\n')
        results = [b.split(":")[1][1:-1] for b in results if "Key Content" in b]
        try:
            list.append(results[0])
        except IndexError:
            list.append("NONE")
    except subprocess.CalledProcessError:
        list.append("NONE")

#write data to excel
kitap = load_workbook("dosya.xlsx")
sheet = kitap.active

for i in range(len(users)):
    sheet.append([users[i],list[i]])

kitap.save("dosya.xlsx")
kitap.close()

        