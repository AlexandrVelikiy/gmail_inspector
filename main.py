import logging
import os
import yaml
import pyzipper
import shutil

from smtplib import SMTP, SMTP_SSL, SMTP_PORT, SMTP_SSL_PORT
from smtplib import SMTPResponseException
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from imap_tools import MailBox, AND
from imap_tools.errors import *

class GmailInspector():
    def __init__(self):
        self.logger = logging.getLogger('gmail_inspector')
        # загружаем конфиг
        self.conf = self.load_config()

    def load_config(self):
        try:
            with open("config.yml", "r") as ymlfile:
                cfg = yaml.load(ymlfile,Loader=yaml.FullLoader)
            return cfg
        except:
            self.logger.exception('load_config')

    def store_email_data(self):
        try:
            find_email = False
            login = self.conf['main']['login']
            password = self.conf['main']['password']
            store_dir = self.conf['other']['tmp_folder']
            prefix_email_subject = self.conf['main']['prefix_email_subject']

            self.logger.info(f'Login to {login}') # логинимся
            try:
                with MailBox('imap.gmail.com').login(login, password) as mailbox:
                    for msg in mailbox.fetch(AND(subject=prefix_email_subject,seen=False)):
                        find_email = True
                        self.logger.info(f'Find email with subjetc contains {prefix_email_subject}')
                        self.logger.info(f'Message subject: {msg.subject}')

                        # создаем папку для сохранения файлов
                        if not os.path.exists(os.path.join(os.getcwd(), store_dir)):
                            os.makedirs(os.path.join(os.getcwd(), store_dir))

                        # сохраняем текст сообщения
                        path = os.path.join(store_dir, 'message_text.txt')
                        self.logger.info(f'Save message body: {path}')
                        with open(path, 'w') as f:
                            f.write(msg.text)

                        # сохраняем вложения
                        self.logger.info(f'Find {len(msg.attachments)} attachments:')
                        for att in msg.attachments:
                            path = os.path.join(store_dir, att.filename)
                            self.logger.info(f'Save attachment: {path}')
                            with open(path,'wb') as f:
                                f.write(att.payload)
                        return True

                    if not find_email:
                        self.logger.info(f'Not found NEW email with subjetc contains {prefix_email_subject}')
                        return False

            except MailboxLoginError:
                self.logger.warning('Login fail')
                return False
        except:
            self.logger.exception('store_email_data')
            return False

    def zip_email_data(self):
        """
        архивируем файлы из каталога store_dir
        и удалем каталог
        """
        try:
            zip_password = self.conf['zip']['password']
            archine_name = self.conf['zip']['archine_name']
            store_dir = self.conf['other']['tmp_folder']

            self.logger.info(f'Compressed  files to {archine_name}')
            with pyzipper.AESZipFile(archine_name, 'w', compression=pyzipper.ZIP_LZMA) as z:
                z.setpassword(zip_password.encode())
                z.setencryption(pyzipper.WZ_AES, nbits=128)
                for root, dirs, files in os.walk(store_dir):
                    for file in files:
                        z.write(os.path.join(root, file))
                        #z.write(file)
            # удаляем папку с сохраненными файлами
            shutil.rmtree(os.path.join(os.getcwd(), store_dir))

        except:
            self.logger.exception('zip_email_data')

    def send_email(self):
        """
        Отправлем сообщение стекстом
        Пароль от архива: archine_password
        И вложением archine_name
        :return:
        """
        try:
            login = self.conf['main']['login']
            password = self.conf['main']['password']
            to_email = self.conf['main']['to_email']
            archine_name = self.conf['zip']['archine_name']
            archine_password = self.conf['zip']['password']

            self.smtp = SMTP_SSL('smtp.gmail.com', SMTP_SSL_PORT)
            self.smtp.login(login, password)

            # создаем письмо
            msg = MIMEMultipart('alternative')
            msg['To'] = to_email
            msg['Subject'] = self.conf['main']['to_email_subject']

            # добваляем в текс сообщения пароль
            msg.attach(MIMEText(f"Пароль от архива: {archine_password}", 'plain'))

            # добваляем во вложение архив
            part = MIMEBase('application', "octet-stream")
            with open(archine_name, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',f'attachment; filename="{archine_name}"')
            msg.attach(part)
            self.logger.info(f'Sending email to {to_email} ...')
            try:
                self.smtp.sendmail(login, to_email,msg.as_bytes())
                self.logger.info('Send email succsesful')
            except SMTPResponseException:
                self.logger.exception('SMTPResponseException')


        except:
            self.logger.exception('zip_email_data')


    def run(self):
        if self.store_email_data():
            self.zip_email_data()
            self.send_email()


def main():
    logger = logging.getLogger('gmail_inspector')
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler('gmail_inspector_log.txt')
    fh.setLevel(logging.DEBUG)

    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    formatter = logging.Formatter('{%(asctime)s} {%(levelname)s} {%(module)s}: "%(message)s"')
    ch.setFormatter(formatter)
    fh.setFormatter(formatter)
    logger.addHandler(ch)
    logger.addHandler(fh)

    gi = GmailInspector()
    gi.run()



if __name__ == '__main__':
    main()

