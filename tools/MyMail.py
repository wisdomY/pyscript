#! /usr/bin/env python
# -*- coding:utf-8 -*-

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header

smtpserver = 'smtp.quncaotech.com'

class MyMail:
    '''This is for email'''
    def _init_(self, user, psw):
        self.userName = user
        self.password = psw

    def Connect(self):
        self.smtp = smtplib.SMTP()
        self.smtp.connect(smtpserver)
        self.smtp.login(self.userName, self.password)

    def DisConnect(self):
        self.smtp.quit()

    def Send(receiver, msg):
        msg = MIMEMultipart('alternative')
        msg['Subject'] = gEmailSubject
        msg['From'] = mailFrom
        msg['To'] = receiver
