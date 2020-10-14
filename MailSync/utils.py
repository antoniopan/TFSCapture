import win32com.client as win32
import time


def send_mail(name_file, mail_header, mail_body, receiver, forworder, img_attachment_list):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    if name_file != '':
        f_name = open(name_file, 'r')
        s_name = f_name.read()
        f_name.close()
        mail.To = s_name
    mail.Recipients.Add('liangliang.pan_HSW-GS')
    if receiver != '':
        mail.Recipients.Add(receiver)
    if forworder != '':
        mail.CC = forworder
    mail.Subject = '%s%s' % (mail_header, time.strftime('%Y-%m-%d', time.localtime()))
    mail.BodyFormat = 2 # html
    i = 1
    for img in img_attachment_list:
        attachment = mail.Attachments.Add(img)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", 'ID00%d' % i)
        i += 1

    mail.HTMLBody = mail_body
    mail.send()
