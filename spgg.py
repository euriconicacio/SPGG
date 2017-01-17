import os
import sys
import time
import socket
import datetime
import webbrowser
from xlrd import open_workbook
from xlwt import easyxf, Workbook
from xlutils.copy import copy
from pyvirtualdisplay import Display
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        Form.setMinimumSize(QtCore.QSize(692, 442))
        Form.setMaximumSize(QtCore.QSize(692, 442))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 691, 441))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.tab)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_3.setObjectName("groupBox_3")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.groupBox_3)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.abrir = QtWidgets.QPushButton(self.groupBox_3)
        self.abrir.setObjectName("abrir")
        self.verticalLayout_3.addWidget(self.abrir)
        self.label_14 = QtWidgets.QLabel(self.groupBox_3)
        self.label_14.setAlignment(QtCore.Qt.AlignCenter)
        self.label_14.setObjectName("label_14")
        self.verticalLayout_3.addWidget(self.label_14)
        self.gridLayout_2.addWidget(self.groupBox_3, 0, 2, 1, 1)
        self.groupBox_5 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_5.setObjectName("groupBox_5")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.groupBox_5)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_resultarq = QtWidgets.QLabel(self.groupBox_5)
        self.label_resultarq.setText("")
        self.label_resultarq.setObjectName("label_resultarq")
        self.verticalLayout_2.addWidget(self.label_resultarq)
        self.abrir_result = QtWidgets.QPushButton(self.groupBox_5)
        self.abrir_result.setMinimumSize(QtCore.QSize(0, 27))
        self.abrir_result.setObjectName("abrir_result")
        self.verticalLayout_2.addWidget(self.abrir_result)
        self.pushButton = QtWidgets.QPushButton(self.groupBox_5)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout_2.addWidget(self.pushButton)
        self.gridLayout_2.addWidget(self.groupBox_5, 9, 2, 1, 1)
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_4.setObjectName("groupBox_4")
        self.gridLayout = QtWidgets.QGridLayout(self.groupBox_4)
        self.gridLayout.setObjectName("gridLayout")
        self.formLayout_3 = QtWidgets.QFormLayout()
        self.formLayout_3.setObjectName("formLayout_3")
        self.label_8 = QtWidgets.QLabel(self.groupBox_4)
        self.label_8.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.label_8.setObjectName("label_8")
        self.formLayout_3.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_8)
        self.label_tempoestimado = QtWidgets.QLabel(self.groupBox_4)
        self.label_tempoestimado.setText("")
        self.label_tempoestimado.setObjectName("label_tempoestimado")
        self.formLayout_3.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.label_tempoestimado)
        self.label_9 = QtWidgets.QLabel(self.groupBox_4)
        self.label_9.setObjectName("label_9")
        self.formLayout_3.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label_9)
        self.label_tempotranscorrido = QtWidgets.QLabel(self.groupBox_4)
        self.label_tempotranscorrido.setText("")
        self.label_tempotranscorrido.setObjectName("label_tempotranscorrido")
        self.formLayout_3.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.label_tempotranscorrido)
        self.label_10 = QtWidgets.QLabel(self.groupBox_4)
        self.label_10.setObjectName("label_10")
        self.formLayout_3.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_10)
        self.progresso = QtWidgets.QProgressBar(self.groupBox_4)
        self.progresso.setProperty("value", 0)
        self.progresso.setObjectName("progresso")
        self.formLayout_3.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.progresso)
        self.label_pontos = QtWidgets.QLabel(self.groupBox_4)
        self.label_pontos.setText("")
        self.label_pontos.setAlignment(QtCore.Qt.AlignCenter)
        self.label_pontos.setObjectName("label_pontos")
        self.formLayout_3.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.label_pontos)
        self.gridLayout.addLayout(self.formLayout_3, 0, 0, 2, 2)
        self.gridLayout_2.addWidget(self.groupBox_4, 9, 0, 1, 2)
        self.gridLayout_3 = QtWidgets.QGridLayout()
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.gerar = QtWidgets.QPushButton(self.tab)
        self.gerar.setObjectName("gerar")
        self.gridLayout_3.addWidget(self.gerar, 0, 1, 1, 1)
        self.limpar = QtWidgets.QPushButton(self.tab)
        self.limpar.setObjectName("limpar")
        self.gridLayout_3.addWidget(self.limpar, 0, 2, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.tab)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout_3.addWidget(self.pushButton_2, 0, 0, 1, 1)
        self.pushButton_3 = QtWidgets.QPushButton(self.tab)
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout_3.addWidget(self.pushButton_3, 0, 3, 1, 1)
        self.gridLayout_2.addLayout(self.gridLayout_3, 7, 0, 1, 3)
        self.groupBox = QtWidgets.QGroupBox(self.tab)
        self.groupBox.setObjectName("groupBox")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.groupBox)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setObjectName("formLayout")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setObjectName("label")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label)
        self.dir = QtWidgets.QComboBox(self.groupBox)
        self.dir.setObjectName("dir")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.dir)
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setObjectName("label_2")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_2)
        self.modelo = QtWidgets.QComboBox(self.groupBox)
        self.modelo.setObjectName("modelo")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.modelo)
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setObjectName("label_3")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.label_3)
        self.func = QtWidgets.QComboBox(self.groupBox)
        self.func.setObjectName("func")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.func)
        self.label_4 = QtWidgets.QLabel(self.groupBox)
        self.label_4.setObjectName("label_4")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.LabelRole, self.label_4)
        self.mare = QtWidgets.QComboBox(self.groupBox)
        self.mare.setObjectName("mare")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.mare)
        self.label_5 = QtWidgets.QLabel(self.groupBox)
        self.label_5.setObjectName("label_5")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.LabelRole, self.label_5)
        self.gzero = QtWidgets.QComboBox(self.groupBox)
        self.gzero.setObjectName("gzero")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.FieldRole, self.gzero)
        self.label_6 = QtWidgets.QLabel(self.groupBox)
        self.label_6.setObjectName("label_6")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.LabelRole, self.label_6)
        self.sisref = QtWidgets.QComboBox(self.groupBox)
        self.sisref.setObjectName("sisref")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.FieldRole, self.sisref)
        self.verticalLayout_4.addLayout(self.formLayout)
        self.gridLayout_2.addWidget(self.groupBox, 0, 0, 4, 2)
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_2.setObjectName("groupBox_2")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.groupBox_2)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.formLayout_2 = QtWidgets.QFormLayout()
        self.formLayout_2.setObjectName("formLayout_2")
        self.label_7 = QtWidgets.QLabel(self.groupBox_2)
        self.label_7.setObjectName("label_7")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_7)
        self.grau = QtWidgets.QSpinBox(self.groupBox_2)
        self.grau.setMaximum(9999)
        self.grau.setProperty("value", 9999)
        self.grau.setObjectName("grau")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.grau)
        self.horizontalLayout.addLayout(self.formLayout_2)
        self.gridLayout_2.addWidget(self.groupBox_2, 2, 2, 1, 1)
        self.tabWidget.addTab(self.tab, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.label_11 = QtWidgets.QLabel(self.tab_3)
        self.label_11.setGeometry(QtCore.QRect(10, 30, 671, 371))
        self.label_11.setText("")
        self.label_11.setPixmap(QtGui.QPixmap("sobre.png"))
        self.label_11.setScaledContents(True)
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(self.tab_3)
        self.label_12.setGeometry(QtCore.QRect(10, 0, 651, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_12.setFont(font)
        self.label_12.setObjectName("label_12")
        self.tabWidget.addTab(self.tab_3, "")

        self.retranslateUi(Form)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Single-Point GEM Generator"))
        self.groupBox_3.setTitle(_translate("Form", "Points Selection"))
        self.abrir.setText(_translate("Form", "Open file"))
        self.label_14.setText(_translate("Form", "There are no points!"))
        self.groupBox_5.setTitle(_translate("Form", "Results"))
        self.abrir_result.setText(_translate("Form", "Open results"))
        self.pushButton.setText(_translate("Form", "Close"))
        self.groupBox_4.setTitle(_translate("Form", "Processing"))
        self.label_8.setText(_translate("Form", "Estimated time:"))
        self.label_9.setText(_translate("Form", "Elapsed time:"))
        self.label_10.setText(_translate("Form", "Progress:"))
        self.gerar.setText(_translate("Form", "Generate extract"))
        self.limpar.setText(_translate("Form", "Clear"))
        self.pushButton_2.setText(_translate("Form", "Test connection"))
        self.pushButton_3.setText(_translate("Form", "Update options"))
        self.groupBox.setTitle(_translate("Form", "Model and Reference Selection"))
        self.label.setText(_translate("Form", "Model Directory:"))
        self.label_2.setText(_translate("Form", "Model File"))
        self.label_3.setText(_translate("Form", "Functional"))
        self.label_4.setText(_translate("Form", "Tide System"))
        self.label_5.setText(_translate("Form", "Zero Degree Term"))
        self.label_6.setText(_translate("Form", "Reference System"))
        self.groupBox_2.setTitle(_translate("Form", "Truncation"))
        self.label_7.setText(_translate("Form", "Maximal Degree"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Form", "Model Generation"))
        self.label_12.setText(_translate("Form", "Help:  http://cienciasgeodesicas.ufpr.br/spgg/en/"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("Form", "About"))

class Ui_Loading(object):
    def setupUi(self, Loading):
        Loading.setObjectName("Loading")
        Loading.resize(308, 53)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Loading.sizePolicy().hasHeightForWidth())
        Loading.setSizePolicy(sizePolicy)
        Loading.setMinimumSize(QtCore.QSize(308, 53))
        Loading.setMaximumSize(QtCore.QSize(308, 53))
        Loading.setCursor(QtGui.QCursor(QtCore.Qt.WaitCursor))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("../icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Loading.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(Loading)
        self.label.setGeometry(QtCore.QRect(10, 10, 281, 31))
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setWordWrap(True)
        self.label.setObjectName("label")

        self.retranslateUi(Loading)
        QtCore.QMetaObject.connectSlotsByName(Loading)

    def retranslateUi(self, Loading):
        _translate = QtCore.QCoreApplication.translate
        Loading.setWindowTitle(_translate("Loading", "Single-Point GEM Generator"))
        self.label.setText(_translate("Loading", "Loading..."))

def verifica_arq():
    if ((fileName == '') or ((fileName[-3:] != 'xls') and (fileName[-4:] != 'xlsx'))):
        alerta("Invalid file format!", "Only xls or xlsx. Please, try again.", 2)
        ui.label_14.setText("There are no points!")
        return False
    else:
        return True

def openFileNameDialog():
    global fileName
    try:
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName()
        if verifica_arq():
            ui.label_14.setText("OK! \n File selected!")
            return fileName
    except:
        pass

def fecha():
    if not os.path.isfile('data.sp0'):
        muda_ext('data.xls')
    raise SystemExit

def alerta(mensagem1, mensagem2,tipo):
        msg = QtWidgets.QMessageBox()
        if tipo == 1:
            msg.setIcon(QtWidgets.QMessageBox.Information)
        else:
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            
        msg.setText(mensagem1)
        msg.setInformativeText(mensagem2)
        msg.setWindowTitle("Single-Point GEM Generator")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        
        retval = msg.exec_()
        return retval

def testa_con():
    try:
        host = socket.gethostbyname("www.google.com")
        s = socket.create_connection((host, 80), 2)
        alerta("Internet Connection OK!", "You have internet connection and you can take advantage of the SPGG features.", 1)
        return True
    except:
        alerta("There is no Internet Connection!", "You don't have internet connection and you can't take advantage of the SPGG features.", 2)
        pass
    return False

def limpa():
    ui.dir.setCurrentIndex(0)
    ui.modelo.setCurrentIndex(0)
    ui.func.setCurrentIndex(0)
    ui.mare.setCurrentIndex(0)
    ui.gzero.setCurrentIndex(0)
    ui.sisref.setCurrentIndex(0)
    ui.label_14.setText("")
    ui.label_tempoestimado.setText("")
    ui.label_pontos.setText("")
    ui.label_resultarq.setText("")
    ui.progresso.setProperty("value", 0)
    if not os.path.isfile('data.sp0'):
        muda_ext('data.xls')
        
def le_combo(nome):
    display = Display(visible=0, size=(800, 600))
    display.start()
    lista=[]

    # Abre site
    browser = webdriver.Chrome()
    browser.get('http://icgem.gfz-potsdam.de/ICGEM/Service.html')    
    
    # Campo 1
    if nome == "model_file":
        campo1 = browser.find_element_by_id("model_directory")
        options1 = [x for x in campo1.find_elements_by_tag_name("option")]
        for element in options1:
            a = str(element.get_attribute("value"))
            print(a)
            campo1s = Select(browser.find_element_by_id('model_directory'))
            campo1s.select_by_visible_text(a)
            campo = browser.find_element_by_id("model_file")
            i = 7
            options = [x for x in campo.find_elements_by_tag_name("option")]    
            if len(options) == 0:
                lista.append(" ")
                lista.append("-")
            for element in options:
                if i == 7:
                    lista.append(" ")
                    i = 9    
                lista.append(str(element.get_attribute("value")))
    else:
        campo = browser.find_element_by_id(nome)
        i = 7
        options = [x for x in campo.find_elements_by_tag_name("option")]    
        if len(options) == 0:
            lista.append(" ")
            lista.append("-")
        for element in options:
            if i == 7:
                lista.append(" ")
                i = 9    
            lista.append(str(element.get_attribute("value")))
    
    return lista

        
def gera_nome_saida(nome_arq):
    if (nome_arq[-3:] == 'xls') or (nome_arq[-3:] == 'odt'):
        a = nome_arq[:-4]
        b = nome_arq[-4:]
    else:
        a = nome_arq[:-5]
        b = nome_arq[-5:]
    return a+'_output'+b

def le_xls(nome_arq, aba):
    wkb=open_workbook(nome_arq)
    sheet=wkb.sheet_by_index(aba)
    
    _matrix=[]
    for row in range (sheet.nrows):
        _row=[]
        for col in range (sheet.ncols):
            _row.append(sheet.cell_value(row,col))
        _matrix.append(_row)
    return _matrix

def preenche_modelos():
    if not os.path.isfile('data.sp0'): muda_ext('data.xls')
    muda_ext('data.sp0')
    nome_arq='data.xls'
    dire=ui.dir.currentText()
    val = []
    ui.modelo.clear()
    ui.modelo.addItem(" ")
    if dire == " ":
        pass
    else:    
        wkb=open_workbook(nome_arq)
        sheet1=wkb.sheet_by_index(0)
        row1 =  []
        i=0
        for row in range (sheet1.nrows):
            row1.append(sheet1.cell_value(row,0))
            if dire == sheet1.cell_value(row,0):
                i = row
        
        sheet2=wkb.sheet_by_index(1)
        row2 = []
        for row in range(sheet2.nrows):
            row2.append(sheet2.cell_value(row,0))
        
        k = 0
        l = 0
        m = 0
        val = []
        while m<len(row2):
            if row2[m] == " ":
                k+=1
            if k == i:
                while l+m<len(row2):
                    if row2[l+m+1] == " ":
                        l = len(row2)
                        k = 666
                        break
                    elif l+m+1 == len(row2)-1:
                        val.append(row2[l+m+1])
                        l = len(row2)
                        k = 666
                        break
                    else:
                        val.append(row2[l+m+1])
                    l+=1
            else:
                m += 1
                pass
            if k==666:
                break
        #print (val)
        #print (len(val))
        ui.modelo.addItems(val)
        muda_ext('data.xls')

def grava_xls(nome_arq,result,aba,tipo):
    if tipo==1:
        rb = open_workbook(nome_arq)
        r_sheet = rb.sheet_by_index(aba)
        wb = copy(rb)
        w_sheet = wb.get_sheet(aba)
    
        numrows = len(result)
        for i in range(numrows):
            w_sheet.write(i, 0, result[i])
    
        wb.save(nome_arq)
    else:
        print (result)
        wb = Workbook()
        ws = wb.add_sheet('Results')
        
        numrows = len(result)
        numcols = len(result[0])
        for i in range(numrows):
            for j in range(numcols):
                ws.write(i, j, result[i][j])
        wb.save(nome_arq)
        
def escreve_lista_arq(nome_arq, a):
    arq = open(nome_arq, "a")
    for item in a:
        arq.write("%s \t" % item)
    arq.write('\n')
    arq.close()

def muda_ext(nome_arq):
    pre, ext = os.path.splitext(nome_arq)
    if ext == ".xls":
        ext1 = ".sp0"
    else:
        ext1 = ".xls"
    os.rename(nome_arq, pre + ext1)    
    
def preenche_combo():
    ui.dir.clear()
    if not os.path.isfile('data.sp0'):
        muda_ext('data.xls')
    muda_ext('data.sp0')
    a = le_xls('data.xls',0)
    for i in range(len(a)):
        ui.dir.addItems(a[i])
    ui.modelo.clear()
    #a = le_xls('data.xls',1)
    #for i in range(len(a)):
    #    ui.modelo.addItems(a[i])
    ui.func.clear()
    a = le_xls('data.xls',2)
    for i in range(len(a)):
        ui.func.addItems(a[i])
    ui.mare.clear()
    a = le_xls('data.xls',3)
    for i in range(len(a)):
        ui.mare.addItems(a[i])
    ui.gzero.clear()
    a = le_xls('data.xls',4)
    for i in range(len(a)):
        ui.gzero.addItems(a[i])
    ui.sisref.clear()
    a = le_xls('data.xls',5)
    for i in range(len(a)):
        ui.sisref.addItems(a[i])
    muda_ext('data.xls')

def conv_tempo(segundos):
    m, s = divmod(segundos, 60)
    h, m = divmod(m, 60)
    d = datetime.time(h,m,s)
    return d
    #ui.label_tempoestimado.setText(str(d))
    

def atualiza_op():
    reply = QtWidgets.QMessageBox()
    reply.setIcon(QtWidgets.QMessageBox.Question)
    reply.setText("Are you sure you want to update the form options?")
    reply.setInformativeText("This operation may take several minutes.")
    reply.setWindowTitle("Single-Point GEM Generator")
    reply.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        
    retval = reply.exec_()
    if retval == QtWidgets.QMessageBox.Yes:
        if(testa_con()):
            Form2.move(350,150)
            Form2.show()
            time.sleep(1)
            muda_ext('data.sp0')
            grava_xls('data.xls',le_combo('model_directory'),0,1)
            grava_xls('data.xls',le_combo('model_file'),1,1)
            grava_xls('data.xls',le_combo('functional'),2,1)
            grava_xls('data.xls',le_combo('tide_system'),3,1)
            grava_xls('data.xls',le_combo('zero_degree_term'),4,1)
            grava_xls('data.xls',le_combo('refsys'),5,1)
            Form2.hide()
            preenche_combo()
            if not os.path.isfile('data.sp0'):
                muda_ext('data.xls')
            alerta("OPTIONS UPDATED!", "The form options were update successfully!", 1)
            return True
        else:
            return False
    else:
        return False
  
def gera_grid_ponto(dire, modelo, func, mare, gzero, sisref, lat, lon, h, grau):
    
    display = Display(visible=0, size=(800, 600))
    display.start()
    
    # Abre site
    browser = webdriver.Chrome()
    browser.get('http://icgem.gfz-potsdam.de/ICGEM/Service.html')
        
    # Model Directory
    model_dir = Select(browser.find_element_by_id('model_directory'))
    model_dir.select_by_visible_text(dire)
    
    # Model File
    model_file = Select(browser.find_element_by_id('model_file'))
    model_file.select_by_visible_text(modelo)
    
    # Functional
    functional = Select(browser.find_element_by_id('functional'))
    #functional.select_by_visible_text('geoid')
    functional.select_by_visible_text(func)
    
    # Tide System
    tide = Select(browser.find_element_by_id('tide_system'))
    #tide.select_by_visible_text('tide_free')
    tide.select_by_visible_text(mare)
    
    # Zero Degree Term
    zerodeg = Select(browser.find_element_by_id('zero_degree_term'))
    #zerodeg.select_by_visible_text('yes')
    zerodeg.select_by_visible_text(gzero)
    
    # Reference System
    ref = Select(browser.find_element_by_id('refsys'))
    #ref.select_by_visible_text('GRS80')
    ref.select_by_visible_text(sisref)
    
    # Grid Step
    step = browser.find_element_by_id('grid_step')
    step.clear()
    step.send_keys('1.0')
    
    # Longitude Limit West
    longlimit_west = browser.find_element_by_id('longlimit_west')
    longlimit_west.clear()
    #longlimit_west.send_keys('-49.2374303888889')
    longlimit_west.send_keys(str(lon))
    
    # Longitude Limit East
    longlimit_east = browser.find_element_by_id('longlimit_east')
    longlimit_east.clear()
    #longlimit_east.send_keys('-49.2374303888889')
    longlimit_east.send_keys(str(lon))
    
    # Latitude Limit South
    latlimit_south = browser.find_element_by_id('latlimit_south')
    latlimit_south.clear()
    #latlimit_south.send_keys('-25.4555288333333')
    latlimit_south.send_keys(str(lat))
    
    # Latitude Limit North
    latlimit_north = browser.find_element_by_id('latlimit_north')
    latlimit_north.clear()
    #latlimit_north.send_keys('-25.4555288333333')
    latlimit_north.send_keys(str(lat))
    
    # Height over Ellipsoid
    hei = browser.find_element_by_id('height_over_ell')
    hei.clear()
    hei.send_keys(str(h))
    
    # Maximal Degree
    max_grau = browser.find_element_by_id('max_used_degree')
    max_grau.clear()
    if grau != '9999':
        max_grau.send_keys(grau)
    else:
        max_grau.send_keys('** max degree of model **')
    
    # ENTER
    start = browser.find_element_by_id('start_but')
    start.send_keys(Keys.ENTER)
    # clica GRID
    
    element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.ID, 'get_but')))
    element.click()
    browser.switch_to_window(browser.window_handles[1])
    try:
        element_present = EC.presence_of_element_located((By.XPATH, '//pre'))
        WebDriverWait(browser, 30).until(element_present)
        a = browser.page_source.encode('utf-8')[-36:-21]
        browser.quit()
        display.stop()
        return float(a.strip())
    except TimeoutException:
        print("Internet issues. Please, try again.")
        return False
    
    
    
def gera_modelos(fileName, dire, modelo, func, mare, gzero, sisref, grau):
    global SAIDA
    i = 0
    pontos = le_xls(fileName,0)
    numrows = len(pontos)
    d = conv_tempo(10*numrows)
    ui.label_tempoestimado.setText(str(d))
    ui.label_pontos.setText("0 of "+str(len(pontos)))
    geo = [[0 for x in range(2)] for y in range(numrows)]
    Form2.move(350,150)
    Form2.show()
    while i < numrows:
        ui.progresso.setValue(i*100/len(pontos))
        ui.label_pontos.setText(str(i+1)+" of "+str(len(pontos)))
        d = conv_tempo(10*(i+1))
        ui.label_tempotranscorrido.setText(str(d))
        geo[i][0] = i+1
        geo[i][1] = gera_grid_ponto(dire,modelo,func,mare,gzero,sisref,pontos[i][1],pontos[i][2],pontos[i][3],grau)
        escreve_lista_arq('temp.txt',geo[i])
        i+=1
    Form2.hide()
    grava_xls(gera_nome_saida(fileName),geo,1,2)
    ui.label_resultarq.setText("File: OK!")
    ui.progresso.setValue(100)
    alerta("Processing finished!", "Click the 'Open results' button to view the results.", 1)
    SAIDA = "ok"
    
def verifica_campos():
    global fileName
    if (ui.dir.currentText() != ' ') and (ui.modelo.currentText() != ' ') and (ui.modelo.currentText() != '-') and (ui.func.currentText() != ' ') and (ui.mare.currentText() != ' ') and (ui.gzero.currentText() != ' ') and (ui.sisref.currentText() != ' ') and (fileName != ''):
        gera_modelos(fileName, ui.dir.currentText(), ui.modelo.currentText(), ui.func.currentText(), ui.mare.currentText(), ui.gzero.currentText(), ui.sisref.currentText(), ui.grau.value())
    else:
        alerta("There are empty fields!", "You must fulfill all the fields to generate the extracts.", 2)
        pass

def abre_resultado():
    global fileName
    global SAIDA
    if SAIDA != "ok":
        alerta("There is no file to open!", "You must run the processing to view the results.", 2)
    else:
        webbrowser.open(gera_nome_saida(fileName))
    

# AQUI COMECA!

if __name__ == "__main__":
    start_time = time.time()
    global fileName
    global SAIDA
    fileName = ''
    SAIDA = ''
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form2 = QtWidgets.QWidget()
    ui2 = Ui_Loading()
    ui2.setupUi(Form2)
    print ("--- %s seconds ---" % (time.time() - start_time))
    if(testa_con()):
        preenche_combo()
        Form.move(300,100)
        Form.show()
        ui.dir.currentIndexChanged.connect(preenche_modelos)
        ui.pushButton_2.clicked.connect(testa_con)
        ui.pushButton.clicked.connect(fecha)
        ui.pushButton_3.clicked.connect(atualiza_op)
        ui.limpar.clicked.connect(limpa)
        ui.abrir.clicked.connect(openFileNameDialog)
        ui.gerar.clicked.connect(verifica_campos)
        ui.abrir_result.clicked.connect(abre_resultado)
    else:
        fecha()
    sys.exit(app.exec_())

