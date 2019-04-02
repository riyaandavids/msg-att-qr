import win32com.client
import os

from pyzbar.pyzbar import decode
from PIL import Image

path = 'c:\\qrcodes\\'

files = [f for f in os.listdir(path)]

for file in files:
    if file.endswith('.msg'):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem(path + file)
        att = msg.Attachments
        for i in att:
            i.SaveAsFile(os.path.join(path, msg.Subject + '.png'))

pngs = [f for f in os.listdir(path)]

for png in pngs:
    if png.endswith('.png'):
        decoded = decode(Image.open(os.path.join(path + png)))
        info = decoded[0].data
        info_d = str(info, 'utf-8')
        f= open(path + 'up.txt',"a+")
        f.write(info_d + "\n")
        f.close()