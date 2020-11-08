import sys
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication
from PyQt5.uic import loadUi
import requests
from requests.auth import HTTPBasicAuth
import getpass
from requests_ntlm import HttpNtlmAuth
from bs4 import BeautifulSoup
import pandas as pd
from time import sleep

arquivo_base = pd.read_excel(r'planilha_numero_sequencial.xlsx')
arquivo_base = pd.read_excel(r'planilha_numero_sequencial.xlsx')
texto_excel = pd.read_excel(r'planilha_textoBase.xlsx')
texto_base = texto_excel['Texto Base'][0]
#lista de expedientes pendentes de resposta (protocolos)
lista_sequenciais = arquivo_base['Seq']
#lista casos pendentes
#lista_sequenciais = ['2811001']


class Login(QDialog):
    def __init__(self):
        super(Login,self).__init__()
        loadUi("login.ui",self)
        self.loginbutton.clicked.connect(self.loginfunction)
        self.password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.createaccbutton.clicked.connect(self.gotocreate)

    def loginfunction(self):
        matricula=self.email.text()
        senha=self.password.text()
        print(matricula)
        usuario = f'corpcaixa\\{matricula}'
        df_base = None
        valorIncrementoProgressBar = 100 / len(arquivo_base)
        for item in lista_sequenciais:
            headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}

            auth=HttpNtlmAuth(usuario, senha)
            session = requests.Session()
            id = item

            response = session.get(f'http://www.portal.dijur.caixa/modulos/subsidios/?pg=subsidiosGravaDoEmail&sequencial={id}', auth=auth, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')
            form = soup.find(id='frmResposta')
            inputs = form.find_all('input')

            dados = {}

            for input in inputs:
                dados[input.get('name')] = input.get('value')

            dados['tipoResposta']='A'
            dados['txaResposta'] = texto_base
            url = 'http://www.portal.dijur.caixa/modulos/subsidios/solicitacao/SubsidioGravaDoEmailGravar.asp'
            r = session.post(url, data=dados)
            print('Status HTTP: ',  r, id)
            print('Sequencial:', id)
            resposta = pd.DataFrame({'Resposta': str(r), 'SEQ': id},index=[0])
            self.progressBar.setValue(valorIncrementoProgressBar)
            valorIncrementoProgressBar += valorIncrementoProgressBar
            if(df_base is None):
                df_base = resposta
            else:
                df_base = pd.concat([df_base, resposta], axis=0, join='outer', ignore_index=True)
            sleep(1)
        df_base.to_csv('resposta.csv', sep=';', index = False, encoding = 'utf-8-sig')
        self.progressBar.setValue(100)
        sleep(2)
        self.close()
        sys.exit(app.exec_())

    def gotocreate(self):
        createacc=CreateAcc()
        widget.addWidget(createacc)
        widget.setCurrentIndex(widget.currentIndex()+1)

class CreateAcc(QDialog):
    def __init__(self):
        super(CreateAcc,self).__init__()
        loadUi("login.ui",self)
        # self.signupbutton.clicked.connect(self.createaccfunction)
        # self.password.setEchoMode(QtWidgets.QLineEdit.Password)
        # self.confirmpass.setEchoMode(QtWidgets.QLineEdit.Password)

    def createaccfunction(self):
        email = self.email.text()
        if self.password.text()==self.confirmpass.text():
            password=self.password.text()
            print("Successfully created acc with email: ", email, "and password: ", password)
            login=Login()
            widget.addWidget(login)
            widget.setCurrentIndex(widget.currentIndex()+1)



app=QApplication(sys.argv)
mainwindow=Login()
widget=QtWidgets.QStackedWidget()
widget.addWidget(mainwindow)
widget.setFixedWidth(480)
widget.setFixedHeight(620)
widget.show()
app.exec_()