# -*- coding: utf-8 -*-
from configparser import ConfigParser
import importlib

import smtplib  # 加载smtplib模块
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class Email:
    def __init__(self) -> None:
        self.cf = ConfigParser()
        import os
        self.confPath = os.getcwd() + '\\config\\config.ini'
        self.cf.read(self.confPath, encoding="utf-8")
        self.MysqlPool = importlib.import_module("config.MySqlPool")
        self.pool = self.MysqlPool.MysqlPool(host="127.0.0.1",
                                             port=3306,
                                             user="user",
                                             password="password",
                                             database="test")
        self.EMAIL_HOST = str(self.isDefual("email.EMAIL_HOST"))
        self.EMAIL_PORT = int(self.isDefual("email.EMAIL_PORT"))
        self.EMAIL_HOST_USER = str(self.isDefual("email.EMAIL_HOST_USER"))
        self.EMAIL_HOST_PASSWORD = str(
            self.isDefual("email.EMAIL_HOST_PASSWORD"))
        self.EMAIL_FROM = str(self.isDefual("email.EMAIL_FROM"))
        self.EMAIL_USE_SSL = bool(self.isDefual("email.EMAIL_USE_SSL"))
        self.EMAIL_USE_TLS = bool(self.isDefual("email.EMAIL_USE_TLS", True))

    def isDefual(self, cf, t=False):
        list_cf = cf.split('.')
        if self.cf[list_cf[0]]['IS_RELOAD'] == "True":
            value = self.pool.fetch_one(
                "select dyNumber from dynamic where dyName = '%s'" %
                list_cf[1])
            self.cf[list_cf[0]][list_cf[1]] = value['dyNumber']
            with open(self.confPath, 'w', encoding='utf-8') as f:
                self.cf.write(f)
            if t:
                self.cf[list_cf[0]]['IS_RELOAD'] = "Flase"
                with open(self.confPath, 'w', encoding='utf-8') as f:
                    self.cf.write(f)
            return value['dyNumber']
        else:
            return self.cf[list_cf[0]][list_cf[1]] if self.cf[list_cf[0]].get(
                list_cf[1]) else self.pool.fetch_one(
                    "select dyNumber from dynamic where dyName = '%s'" %
                    list_cf[1])['dyNumber']

    def send_email(self, file_list, addressee, email_text):
        try:
            # 创建一个带附件的实例
            msg = MIMEMultipart()
            # 发件人格式
            msg['From'] = formataddr([self.EMAIL_FROM, self.EMAIL_HOST_USER])
            # 收件人格式
            msg['To'] = addressee
            # 邮件主题
            msg['Subject'] = email_text

            # 邮件正文内容
            msg.attach(MIMEText(email_text, 'plain', 'utf-8'))
            # 多个附件
            for file_name in file_list:
                # 构造附件
                xlsxpart = MIMEApplication(open(file_name, 'rb').read())
                # filename表示邮件中显示的附件名
                xlsxpart.add_header('Content-Disposition',
                                    'attachment',
                                    filename='%s' % file_name.split('\\')[-1])
                msg.attach(xlsxpart)

            # SMTP服务器
            server = smtplib.SMTP_SSL(self.EMAIL_HOST,
                                      self.EMAIL_PORT,
                                      timeout=1000)
            # 登录账户
            server.login(self.EMAIL_HOST_USER, self.EMAIL_HOST_PASSWORD)
            # 发送邮件
            server.sendmail(self.EMAIL_HOST_USER, [addressee], msg.as_string())
            # 退出账户
            server.quit()
            return True

        except Exception as e:
            print(e)
            return False
