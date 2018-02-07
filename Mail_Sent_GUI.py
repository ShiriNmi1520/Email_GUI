from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename
import ctypes
import smtplib

root = Tk()
root.title("公告信發送")
root.resizable(False, False)

receiver_mail = []


def mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def send():
    msg = MIMEMultipart()
    msg["Subject"] = mail_title.get()
    msg["From"] = email_send.get()
    msg["To"] = ', '.join(receiver_mail)
    msg.attach(part)
    try:
        global smtp
        smtp = smtplib.SMTP("SMTP_Ip_Addr", "SMTP_Port")  # smtp伺服器位址，採TLS連線
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login("Account", "Password")
        # 登入SMTP使用的Account, Password皆用同一組，僅需修改發信email_addr
        smtp.sendmail(email_send.get(), receiver_mail, msg.as_string())
        mbox("寄信", "信件已成功寄出", 0)
    except smtplib.SMTPSenderRefused:
        mbox("錯誤", "此寄信位址被SMTP伺服器拒絕", 0)
    except smtplib.SMTPDataError:
        mbox("錯誤", "SMTP伺服器拒絕送信", 0)
    except smtplib.SMTPRecipientsRefused:
        mbox("錯誤", "收信位址拒絕收信", 0)
    except smtplib.SMTPConnectError:
        mbox("錯誤", "嘗試與SMTP伺服器連線時發生錯誤", 0)
    except smtplib.SMTPServerDisconnected:
        mbox("錯誤", "與SMTP伺服器間連線發生未預期的斷線", 0)
    except smtplib.SMTPHeloError:
        mbox("錯誤", "與SMTP建立初次連線時發生錯誤", 0)
    except smtplib.SMTPAuthenticationError:
        mbox("錯誤", "SMTP拒絕登入(帳號密碼錯誤)", 0)
    except:
        mbox("錯誤", "信件未成功寄出", 0)
    finally:
        smtp.close()


def choose_file():
    global file
    global part
    file = askopenfilename(title="Choose File")
    part = MIMEApplication(open(file, "rb").read())
    part.add_header("Content-Disposition",
                    "attachment", filename=basename(file))
    mbox("已選擇信件檔", "信件檔的位置在%s" % file, 0)


def choose_receiver():
    global receiver
    receiver = askopenfilename(title="Choose List For Receiver")
    try:
        fp = open(str(receiver), "r")
        for line in fp:
            receiver_mail.append(line)
        mbox("已選擇收信者清單", "清單的位置在%s" % receiver, 0)
    except FileNotFoundError:
        mbox("錯誤", "未選擇檔案或檔案位置錯誤!", 0)
    except:
        mbox("錯誤", "無法開啟收件者清單", 0)


Label(root, text="標題:").grid(row=0, sticky=W)
Label(root, text="寄件者:").grid(row=1, sticky=W)
Label(root, text="寄件者信箱:").grid(row=2, sticky=W)

mail_title = Entry(width=40)
sender = Entry(width=40)
email_send = Entry(width=40)

mail_title.grid(row=0, column=1, sticky=W)
sender.grid(row=1, column=1, sticky=W)
email_send.grid(row=2, column=1, sticky=W)

Button(root, text='退出', command=root.quit).grid(row=6, column=0)
Button(root, text="發送", command=send).grid(row=6, column=1)
Button(root, text="選擇信件檔", command=choose_file).grid(row=7, column=0)
Button(root, text="選擇收件者清單", command=choose_receiver).grid(row=7, column=1)

mainloop()
