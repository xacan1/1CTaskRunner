#Author:Smirnov Hasan 2020

import sys
import win32com.client
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib, ssl

class EmailSender:
    def __init__(self, login, password, receiver_emails, subject):
        self.port = 465
        self.smtp_server = 'smtp.yandex.ru'
        self.login = login
        self.password = password
        self.subject = subject
        self.receiver_emails = receiver_emails
        self.context = ssl.create_default_context()# Создание безопасного контекста SSL 

    def SendMessage(self, message):
        msg = MIMEMultipart()
        msg['From'] = self.login
        msg['To'] = self.receiver_emails
        msg['Subject'] = self.subject
        msg.attach(MIMEText(message, 'plain'))
        
        with smtplib.SMTP_SSL(self.smtp_server, self.port, context=self.context) as server_mail:
            server_mail.login(self.login, self.password)
            server_mail.sendmail(msg['From'], msg['To'], msg.as_string())
            server_mail.quit()

class Exchanger1C:
    def __init__(self, path=''):
        self.path = path
        if self.path:
            self.path = self.path + '\\'
        self.parameters = self.GetParameters()
        self.v83com = None
        self.email_sender = None

    def GetParameters(self):
        parameters = {}
        
        try:
            fparams = open(self.path + 'parameters.txt', 'r')
        except Exception as exp:
            self.Logging(exp)
            return parameters

        parameters['TEST'] = 'False'
        
        for fline in fparams:
            if fline[0] == '#':
                continue
            line = fline.strip().split(':')
            parameters[line[0].strip()] = line[1].strip()

        fparams.close()
        return parameters

    #отправим уведомление по почте и запишем в файл
    def Logging(self, text_error):
        current_time = time.strftime("%d.%m.%Y %H:%M:%S", time.localtime())
        msg = current_time + ' | ' + text_error
        
        with open(self.path + 'errors.log', 'a') as flog:
            flog.write(msg + '\n')
        
        if len(self.parameters['MAIL']) > 5 and len(self.parameters['PASS']) > 3 and len(self.parameters['ADDR']) > 5 and len(self.parameters['SUBJ']) > 1:
            postman = EmailSender(self.parameters['MAIL'], self.parameters['PASS'], self.parameters['ADDR'], self.parameters['SUBJ'])
            postman.SendMessage(msg)

    def TestConnect(self):
        self.v83com = self.GetConnectTo1C()
        print(self.GetCode())
        
        if self.v83com:
            self.Logging('TEST EMAIL 1C: OK')
            print('Соединение успешно установлено')
            self.v83com = None
        else:
            self.Logging('TEST EMAIL 1C: Не удалось установить соединение')
            print('Не удалось установить соединение')
    
    def GetConnectTo1C(self):
        v83com = None
        
        if not self.parameters:
            return v83com
        
        try:
            v83com = win32com.client.Dispatch(self.parameters['VERS'] + '.COMConnector').Connect(self.parameters['CSTR'])
        except Exception as exp:
            self.Logging(exp)
            v83com = None
            return v83com

        return v83com

    def GetCode(self):
        code = ''
        
        try:
            fcode = open(self.path + 'code.txt', 'r')
            code = fcode.read()
            fcode.close()
        except Exception as exp:
            self.Logging(exp)
            return code
            
        return code

    def StartProcedureFrom1C(self):
        if self.parameters and self.parameters['TEST'] == 'True':
            self.TestConnect()
            return
        
        self.v83com = self.GetConnectTo1C()
        
        if not self.v83com:
            return
        
        code = self.GetCode()

        if len(code) < 3:
            self.Logging('Не найден код для выполнения')
            return
        
        try:
            returnValue = self.v83com.ПроцедурыВнешнегоСоединения.ExecutionOfExternalCode(code)
        except Exception as exp:
            self.Logging(exp)
            self.v83com = None

##    #Эта функция просто для примера как можно получать данные из базы
##    def GetDataFrom1C(self):
##        self.v83com = self.GetConnectTo1C()
##        
##        if not self.v83com:
##            return
##        
##        select_docs = self.v83com.Documents.СчетНаОплатуПокупателю.Select()
##        
##        while select_docs.Next():
##            Number = select_docs.Номер
##            SumDoc = select_docs.СуммаДокумента
##            print('Номер: {} сумма: {} руб'.format(Number, SumDoc))

if __name__ == '__main__':
    
    if len(sys.argv) > 1:
        exchange = Exchanger1C(sys.argv[1]) #при запуске из планировщика программа думает, что ее запустили из папки Windows и ищет файл настройки там, по этому надо передать путь к ней через параметр
    else:
        exchange = Exchanger1C()
        
    exchange.StartProcedureFrom1C()
