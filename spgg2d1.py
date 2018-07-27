import os
import sys
import time
import glob
import socket
import datetime
import webbrowser
import collections
import urllib.request
import json 
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
        Form.resize(1000, 680)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        Form.setMinimumSize(QtCore.QSize(1000, 680))
        Form.setMaximumSize(QtCore.QSize(1000, 680))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Form.setWindowIcon(icon)
        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 1001, 681))
        self.tabWidget.setAutoFillBackground(True)
        self.tabWidget.setStyleSheet("QTabWidget::pane { /* The tab widget frame */\n"
"    border-top: 2px solid #C2C7CB;\n"
"    background-color: #a9cfea;\n"
"}\n"
"\n"
"QTabWidget::tab-bar {\n"
"    left: 5px; /* move to the right by 5px */\n"
"}\n"
"\n"
"/* Style the tab using the tab sub-control. Note that\n"
"    it reads QTabBar _not_ QTabWidget */\n"
"QTabBar::tab {\n"
"    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,\n"
"                                stop: 0 #E1E1E1, stop: 0.4 #DDDDDD,\n"
"                                stop: 0.5 #D8D8D8, stop: 1.0 #D3D3D3);\n"
"    border: 2px solid #C4C4C3;\n"
"    border-bottom-color: #C2C7CB; /* same as the pane color */\n"
"    border-top-left-radius: 4px;\n"
"    border-top-right-radius: 4px;\n"
"    min-width: 8ex;\n"
"    padding: 2px;\n"
"}\n"
"\n"
"QTabBar::tab:selected, QTabBar::tab:hover {\n"
"    background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,\n"
"                                stop: 0 #fafafa, stop: 0.4 #f4f4f4,\n"
"                                stop: 0.5 #e7e7e7, stop: 1.0 #fafafa);\n"
"}\n"
"\n"
"QTabBar::tab:selected {\n"
"    border-color: #9B9B9B;\n"
"    border-bottom-color: #C2C7CB; /* same as pane color */\n"
"}\n"
"\n"
"QTabBar::tab:!selected {\n"
"    margin-top: 2px; /* make non-selected tabs look smaller */\n"
"}")
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_3.setGeometry(QtCore.QRect(810, 10, 181, 101))
        self.groupBox_3.setStyleSheet("\n"
"\n"
"QGroupBox {\n"
"    border: 1px solid gray;\n"
"    border-radius: 9px;\n"
"    margin-top: 0.5em;\n"
"    background-color: #dfeaf4;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QGroupBox::title {\n"
"    subcontrol-origin: margin;\n"
"    left: 10px;\n"
"    padding: 0 3px 0 3px;\n"
"    background-color: #dfeaf4;\n"
"    border-radius: 9px;\n"
"}\n"
"\n"
"")
        self.groupBox_3.setObjectName("groupBox_3")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.groupBox_3)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.abrir = QtWidgets.QPushButton(self.groupBox_3)
        self.abrir.setObjectName("abrir")
        self.verticalLayout_2.addWidget(self.abrir)
        self.label_14 = QtWidgets.QLabel(self.groupBox_3)
        self.label_14.setAlignment(QtCore.Qt.AlignCenter)
        self.label_14.setObjectName("label_14")
        self.verticalLayout_2.addWidget(self.label_14)
        self.groupBox = QtWidgets.QGroupBox(self.tab)
        self.groupBox.setGeometry(QtCore.QRect(10, 10, 791, 501))
        self.groupBox.setStyleSheet("\n"
"\n"
"QGroupBox {\n"
"    border: 1px solid gray;\n"
"    border-radius: 9px;\n"
"    margin-top: 0.5em;\n"
"    background-color: #dfeaf4;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QGroupBox::title {\n"
"    subcontrol-origin: margin;\n"
"    left: 10px;\n"
"    padding: 0 3px 0 3px;\n"
"    border-radius: 9px;\n"
"    background-color: #dfeaf4;\n"
"}\n"
"\n"
"")
        self.groupBox.setObjectName("groupBox")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.groupBox)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.formLayout_2 = QtWidgets.QFormLayout()
        self.formLayout_2.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
        self.formLayout_2.setObjectName("formLayout_2")
        self.label_4 = QtWidgets.QLabel(self.groupBox)
        self.label_4.setObjectName("label_4")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_4)
        self.ref_sys = QtWidgets.QComboBox(self.groupBox)
        self.ref_sys.setObjectName("ref_sys")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.ref_sys)
        self.frame = QtWidgets.QFrame(self.groupBox)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.formLayout_2.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.frame)
        self.label_5 = QtWidgets.QLabel(self.groupBox)
        self.label_5.setObjectName("label_5")
        self.formLayout_2.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.label_5)
        self.tide_sys = QtWidgets.QComboBox(self.groupBox)
        self.tide_sys.setObjectName("tide_sys")
        self.formLayout_2.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.tide_sys)
        self.label_6 = QtWidgets.QLabel(self.groupBox)
        self.label_6.setObjectName("label_6")
        self.formLayout_2.setWidget(4, QtWidgets.QFormLayout.LabelRole, self.label_6)
        self.zero_deg = QtWidgets.QComboBox(self.groupBox)
        self.zero_deg.setObjectName("zero_deg")
        self.formLayout_2.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.zero_deg)
        self.formLayout_5 = QtWidgets.QFormLayout()
        self.formLayout_5.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
        self.formLayout_5.setObjectName("formLayout_5")
        self.radiusLabel = QtWidgets.QLabel(self.groupBox)
        self.radiusLabel.setObjectName("radiusLabel")
        self.formLayout_5.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.radiusLabel)
        self.radius = QtWidgets.QLineEdit(self.groupBox)
        self.radius.setEnabled(False)
        self.radius.setObjectName("radius")
        self.formLayout_5.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.radius)
        self.gMLabel = QtWidgets.QLabel(self.groupBox)
        self.gMLabel.setObjectName("gMLabel")
        self.formLayout_5.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.gMLabel)
        self.gm = QtWidgets.QLineEdit(self.groupBox)
        self.gm.setEnabled(False)
        self.gm.setObjectName("gm")
        self.formLayout_5.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.gm)
        self.flatLabel = QtWidgets.QLabel(self.groupBox)
        self.flatLabel.setObjectName("flatLabel")
        self.formLayout_5.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.flatLabel)
        self.flat = QtWidgets.QLineEdit(self.groupBox)
        self.flat.setEnabled(False)
        self.flat.setObjectName("flat")
        self.formLayout_5.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.flat)
        self.omegaLabel = QtWidgets.QLabel(self.groupBox)
        self.omegaLabel.setObjectName("omegaLabel")
        self.formLayout_5.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.omegaLabel)
        self.omega = QtWidgets.QLineEdit(self.groupBox)
        self.omega.setEnabled(False)
        self.omega.setObjectName("omega")
        self.formLayout_5.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.omega)
        self.formLayout_2.setLayout(2, QtWidgets.QFormLayout.FieldRole, self.formLayout_5)
        self.gridLayout_4.addLayout(self.formLayout_2, 3, 0, 1, 4)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setStyleSheet("border: none;")
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setStyleSheet("border: none;")
        self.label_3.setObjectName("label_3")
        self.horizontalLayout.addWidget(self.label_3)
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setStyleSheet("border: none;")
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.label_15 = QtWidgets.QLabel(self.groupBox)
        self.label_15.setStyleSheet("border: none;")
        self.label_15.setObjectName("label_15")
        self.horizontalLayout.addWidget(self.label_15)
        self.gridLayout_4.addLayout(self.horizontalLayout, 0, 0, 1, 4)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.dir = QtWidgets.QListWidget(self.groupBox)
        self.dir.setStyleSheet("border: 1px solid gray;")
        self.dir.setObjectName("dir")
        self.horizontalLayout_2.addWidget(self.dir)
        self.modelo = QtWidgets.QListWidget(self.groupBox)
        self.modelo.setStyleSheet("border: 1px solid gray;")
        self.modelo.setObjectName("modelo")
        self.horizontalLayout_2.addWidget(self.modelo)
        self.other = QtWidgets.QListWidget(self.groupBox)
        self.other.setStyleSheet("border: 1px solid gray;")
        self.other.setObjectName("other")
        self.horizontalLayout_2.addWidget(self.other)
        self.func = QtWidgets.QListWidget(self.groupBox)
        self.func.setStyleSheet("border: 1px solid gray;")
        self.func.setObjectName("func")
        self.horizontalLayout_2.addWidget(self.func)
        self.gridLayout_4.addLayout(self.horizontalLayout_2, 1, 0, 1, 4)
        self.groupBox_5 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_5.setGeometry(QtCore.QRect(810, 520, 181, 121))
        self.groupBox_5.setStyleSheet("\n"
"\n"
"QGroupBox {\n"
"    border: 1px solid gray;\n"
"    border-radius: 9px;\n"
"    margin-top: 0.5em;\n"
"    background-color: #dfeaf4;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QGroupBox::title {\n"
"    subcontrol-origin: margin;\n"
"    left: 10px;\n"
"    padding: 0 3px 0 3px;\n"
"    background-color: #dfeaf4;\n"
"    border-radius: 9px;\n"
"}\n"
"\n"
"")
        self.groupBox_5.setObjectName("groupBox_5")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.groupBox_5)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.abrir_result = QtWidgets.QPushButton(self.groupBox_5)
        self.abrir_result.setMinimumSize(QtCore.QSize(0, 27))
        self.abrir_result.setObjectName("abrir_result")
        self.verticalLayout_4.addWidget(self.abrir_result)
        self.fechar = QtWidgets.QPushButton(self.groupBox_5)
        self.fechar.setObjectName("fechar")
        self.verticalLayout_4.addWidget(self.fechar)
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_4.setGeometry(QtCore.QRect(10, 520, 791, 121))
        self.groupBox_4.setStyleSheet("\n"
"\n"
"QGroupBox {\n"
"    border: 1px solid gray;\n"
"    border-radius: 9px;\n"
"    margin-top: 0.5em;\n"
"    background-color: #dfeaf4;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QGroupBox::title {\n"
"    subcontrol-origin: margin;\n"
"    left: 10px;\n"
"    padding: 0 3px 0 3px;\n"
"    background-color: #dfeaf4;\n"
"    border-radius: 9px;\n"
"}\n"
"\n"
"")
        self.groupBox_4.setObjectName("groupBox_4")
        self.formLayout_4 = QtWidgets.QFormLayout(self.groupBox_4)
        self.formLayout_4.setObjectName("formLayout_4")
        self.label_8 = QtWidgets.QLabel(self.groupBox_4)
        self.label_8.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.label_8.setObjectName("label_8")
        self.formLayout_4.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_8)
        self.label_tempoestimado = QtWidgets.QLabel(self.groupBox_4)
        self.label_tempoestimado.setText("")
        self.label_tempoestimado.setObjectName("label_tempoestimado")
        self.formLayout_4.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.label_tempoestimado)
        self.label_9 = QtWidgets.QLabel(self.groupBox_4)
        self.label_9.setObjectName("label_9")
        self.formLayout_4.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label_9)
        self.label_tempotranscorrido = QtWidgets.QLabel(self.groupBox_4)
        self.label_tempotranscorrido.setText("")
        self.label_tempotranscorrido.setObjectName("label_tempotranscorrido")
        self.formLayout_4.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.label_tempotranscorrido)
        self.label_10 = QtWidgets.QLabel(self.groupBox_4)
        self.label_10.setObjectName("label_10")
        self.formLayout_4.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_10)
        self.progresso = QtWidgets.QProgressBar(self.groupBox_4)
        self.progresso.setMaximumSize(QtCore.QSize(16777215, 21))
        self.progresso.setProperty("value", 0)
        self.progresso.setObjectName("progresso")
        self.formLayout_4.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.progresso)
        self.label_pontos = QtWidgets.QLabel(self.groupBox_4)
        self.label_pontos.setText("")
        self.label_pontos.setAlignment(QtCore.Qt.AlignCenter)
        self.label_pontos.setObjectName("label_pontos")
        self.formLayout_4.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.label_pontos)
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_2.setGeometry(QtCore.QRect(810, 140, 181, 111))
        self.groupBox_2.setStyleSheet("\n"
"\n"
"QGroupBox {\n"
"    border: 1px solid gray;\n"
"    border-radius: 9px;\n"
"    margin-top: 0.5em;\n"
"    background-color: #dfeaf4;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QGroupBox::title {\n"
"    subcontrol-origin: margin;\n"
"    left: 10px;\n"
"    padding: 0 3px 0 3px;\n"
"    background-color: #dfeaf4;\n"
"    border-radius: 9px;\n"
"}\n"
"\n"
"")
        self.groupBox_2.setObjectName("groupBox_2")
        self.formLayout = QtWidgets.QFormLayout(self.groupBox_2)
        self.formLayout.setObjectName("formLayout")
        self.label_13 = QtWidgets.QLabel(self.groupBox_2)
        self.label_13.setObjectName("label_13")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label_13)
        self.label_7 = QtWidgets.QLabel(self.groupBox_2)
        self.label_7.setObjectName("label_7")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_7)
        self.grau = QtWidgets.QSpinBox(self.groupBox_2)
        self.grau.setMaximum(9999)
        self.grau.setProperty("value", 9999)
        self.grau.setObjectName("grau")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.grau)
        self.cut = QtWidgets.QSpinBox(self.groupBox_2)
        self.cut.setMaximum(9999)
        self.cut.setProperty("value", 9999)
        self.cut.setObjectName("cut")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.cut)
        self.groupBox_6 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_6.setGeometry(QtCore.QRect(810, 280, 181, 211))
        self.groupBox_6.setStyleSheet("\n"
"\n"
"QGroupBox {\n"
"    border: 1px solid gray;\n"
"    border-radius: 9px;\n"
"    margin-top: 0.5em;\n"
"    background-color: #dfeaf4;\n"
"    font-weight: bold;\n"
"}\n"
"\n"
"QGroupBox::title {\n"
"    subcontrol-origin: margin;\n"
"    left: 10px;\n"
"    padding: 0 3px 0 3px;\n"
"    background-color: #dfeaf4;\n"
"    border-radius: 9px;\n"
"}\n"
"\n"
"")
        self.groupBox_6.setObjectName("groupBox_6")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.groupBox_6)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 30, 161, 171))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.pushButton_2 = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout.addWidget(self.pushButton_2)
        self.limpar = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.limpar.setObjectName("limpar")
        self.verticalLayout.addWidget(self.limpar)
        self.gerar = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.gerar.setObjectName("gerar")
        self.verticalLayout.addWidget(self.gerar)
        self.tabWidget.addTab(self.tab, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.label_11 = QtWidgets.QLabel(self.tab_3)
        self.label_11.setGeometry(QtCore.QRect(10, 60, 971, 561))
        self.label_11.setText("")
        self.label_11.setPixmap(QtGui.QPixmap("sobre.png"))
        self.label_11.setScaledContents(True)
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(self.tab_3)
        self.label_12.setGeometry(QtCore.QRect(270, 10, 470, 26))
        font = QtGui.QFont()
        font.setPointSize(14)
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
        Form.setWindowTitle(_translate("Form", "Single-Point GEM Generator - v 2.1"))
        self.groupBox_3.setTitle(_translate("Form", "Points selection"))
        self.abrir.setText(_translate("Form", "Open..."))
        self.label_14.setText(_translate("Form", "No points loaded."))
        self.groupBox.setTitle(_translate("Form", "Model and functional parameters"))
        self.label_4.setText(_translate("Form", "Reference System"))
        self.label_5.setText(_translate("Form", "Tide System"))
        self.label_6.setText(_translate("Form", "Zero Degree Term"))
        self.radiusLabel.setText(_translate("Form", "Radius"))
        self.gMLabel.setText(_translate("Form", "GM"))
        self.flatLabel.setText(_translate("Form", "Flat"))
        self.omegaLabel.setText(_translate("Form", "Omega"))
        self.label.setText(_translate("Form", "Directory"))
        self.label_3.setText(_translate("Form", "Model"))
        self.label_2.setText(_translate("Form", "Other options"))
        self.label_15.setText(_translate("Form", "Functional"))
        self.groupBox_5.setTitle(_translate("Form", "Results"))
        self.abrir_result.setText(_translate("Form", "Open results"))
        self.fechar.setText(_translate("Form", "Close"))
        self.groupBox_4.setTitle(_translate("Form", "Processing"))
        self.label_8.setText(_translate("Form", "Estimated Time:"))
        self.label_9.setText(_translate("Form", "Elapsed Time:"))
        self.label_10.setText(_translate("Form", "Progress:"))
        self.groupBox_2.setTitle(_translate("Form", "Truncation"))
        self.label_13.setText(_translate("Form", "Gentle cut"))
        self.label_7.setText(_translate("Form", "Max degree"))
        self.groupBox_6.setTitle(_translate("Form", "Calculation"))
        self.pushButton_2.setText(_translate("Form", "Test Connection"))
        self.limpar.setText(_translate("Form", "Clear"))
        self.gerar.setText(_translate("Form", "Generate extracts"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Form", "Extract Generation"))
        self.label_12.setText(_translate("Form", "Help:  http://euriconicacio.github.io/spgg/"))
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
        icon.addPixmap(QtGui.QPixmap("icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
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
        self.label.setText(_translate("Loading", "Loading... Wait, this may take several minutes..."))

        
ref_sys = {"WGS84":["6378137.0", "298.257223563", "3.986004418e+14", "7.292115e-5"], "GRS80":["6378137.0", "298.257222101", "3.986005e+14", "7.292115e-5"], "GRIM5":["6378136.46", "298.25765", "3.986004415e+14", "7.292115e-5"], "EGM2008":[ "6378136.58", "298.257686", "3.986004415e+14", "7.292115e-5"], "GRS67":[ "6378160.0", "298.247167427", "3.986030e+14", "7.2921151467e-5"], "JGM3":["6378136.3",  "298.25700", "3.986004415e+14", "7.29211500e-5" ], "WGS72":["6378135.0", "298.257", "3.986005e+14", "7.292115e-5"], "NIMA96":[ "6378136.0", "298.256415099", "3.986004415e+14", "7.292115e-5" ], "MOON":[ "1738140", "3240", "4.902801076e+12", "2.661621e-6" ], "MARS":[ "3397000", "191.18", "4.2828369773938997e+13", "7.08822e-5" ], "VENUS":[ "6051000", "150700", "3.248585897260e+14", "2.9926e-7"]}
ref_sys = collections.OrderedDict(ref_sys)

refsys_set = {"Earth":[ "WGS84", "GRS80", "GRIM5", "EGM2008", "GRS67", "JGM3", "WGS72", "NIMA96", "- user defined -"], "Moon": [ "MOON", "- user defined -"], "Mars": [ "MARS", "- user defined -"], "Venus":[ "VENUS", "- user defined -"]}
refsys_set = collections.OrderedDict(refsys_set)

tipos = [("Longtime Model",'longtime'),("Model from Series",'series'),("Topography related Model",'reltopo'),("Celestial Object Model",'celestial'),("Topography",'topo')]
tipos = collections.OrderedDict(tipos)

funcionais = {'longtime':['height_anomaly', 'height_anomaly_ell', 'geoid', 'gravity_disturbance', 'gravity_disturbance_geoid', 'gravity_disturbance_sa', 'gravity_anomaly', 'gravity_anomaly_cl', 'gravity_anomaly_sa', 'gravity_anomaly_bg', 'gravity_earth', 'gravity_ell', 'potential_ell', 'gravitation_ell', 'second_r_derivative', 'water_column'], 'series': ['height_anomaly', 'height_anomaly_ell', 'geoid', 'gravity_disturbance', 'gravity_disturbance_geoid', 'gravity_disturbance_sa', 'gravity_anomaly', 'gravity_anomaly_cl', 'gravity_anomaly_sa', 'gravity_anomaly_bg', 'gravity_earth', 'gravity_ell', 'potential_ell', 'gravitation_ell', 'second_r_derivative', 'water_column'], 'reltopo':['height_anomaly', 'height_anomaly_ell', 'geoid', 'gravity_disturbance', 'gravity_disturbance_geoid', 'gravity_disturbance_sa', 'gravity_anomaly', 'gravity_anomaly_cl', 'gravity_anomaly_sa', 'gravity_anomaly_bg', 'gravity_earth', 'gravity_ell', 'potential_ell', 'gravitation_ell', 'second_r_derivative', 'water_column'], 'celestial':['height_anomaly', 'height_anomaly_ell', 'geoid', 'gravity_disturbance', 'gravity_disturbance_geoid', 'gravity_disturbance_sa', 'gravity_anomaly', 'gravity_anomaly_cl', 'gravity_anomaly_sa', 'gravity_anomaly_bg', 'gravity_earth', 'gravity_ell', 'potential_ell', 'gravitation_ell', 'second_r_derivative', 'water_column'], 'topo':['topography_shm', 'topography_grd']}

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█'):
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = '\r')
    if iteration == total: 
        print()

def verifica_arq():
    if (fileName == '') or (fileName[-3:] != 'xls'):
        alerta("Invalid file format!", "Only xls. Please, try again.", 2)
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
    raise SystemExit

def alerta(mensagem1, mensagem2,tipo):
        msg = QtWidgets.QMessageBox()
        if tipo == 1:
            msg.setIcon(QtWidgets.QMessageBox.Information)
        else:
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            
        msg.setText(mensagem1)
        msg.setInformativeText(mensagem2)
        msg.setWindowTitle("Single-Point GEM Generator - v 2.1")
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
    ui.dir.clear()
    for value,key in tipos.items():
        ui.dir.addItem(value)
    ui.modelo.clear()
    ui.other.clear()
    ui.func.clear()
    ui.tide_sys.setCurrentIndex(0)
    ui.zero_deg.setCurrentIndex(0)
    ui.ref_sys.setCurrentIndex(0)
    ui.radius.setText("")
    ui.gm.setText("")
    ui.flat.setText("")
    ui.omega.setText("")
    ui.radius.setEnabled(False)
    ui.gm.setEnabled(False)
    ui.flat.setEnabled(False)
    ui.omega.setEnabled(False)
    ui.cut.setValue(9999)
    ui.grau.setValue(9999)
    ui.label_14.setText("No points loaded.")
    ui.label_tempoestimado.setText("")
    ui.label_tempotranscorrido.setText("")
    ui.label_pontos.setText("")
    ui.progresso.setProperty("value", 0)
        
def gera_nome_saida(nome_arq):
    if (nome_arq[-3:] == 'xls'):
        a = nome_arq[:-4]
        b = nome_arq[-4:]
    else:
        a = nome_arq[:-5]
        b = nome_arq[-4:]
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
        #print (result)
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


def conv_tempo(segundos):
    t = datetime.timedelta(seconds=segundos);
    resp = "{}d {} (dd hh:mm:ss)".format(t.days%30, datetime.timedelta(seconds=t.seconds));
    return resp
    
def gera_grid_ponto(dire, other, modelo, func, mare, lat, long, hovell, sisref, raio, const_gm, achat, velrot, gzero, gentlecut, graumax):
    currfolder = os.getcwd()
    display = Display(visible=0, size=(1500, 1000))
    display.start()

    # Abre site com configuracoes
    chromeoptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : currfolder, "extensions_to_open": "gdf","profile.managed_default_content_settings.images":2}
    chromeoptions.add_experimental_option("prefs",prefs)
    browser = webdriver.Chrome(chrome_options=chromeoptions)

    browser.get('http://icgem.gfz-potsdam.de/calc')

    # Model Directory
    model_dir = Select(browser.find_element_by_name('sel_type'))
    model_dir.select_by_visible_text(dire)

    if dire == "Model from Series":
        # Series
        browser.implicitly_wait(10)
        element = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element((By.XPATH, "//select[@name='sel_series']"),other))
        time.sleep(1)
        model_other = Select(browser.find_element_by_name('sel_series'))
        model_other.select_by_visible_text(other)
        browser.execute_script("series_change()")

        # Model
        element = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element((By.XPATH, "//select[@name='sel_model']"),modelo))
        time.sleep(1)
        model_file = Select(browser.find_element_by_name('sel_model'))
        model_file.select_by_visible_text(modelo)
        browser.execute_script("model_change()")

    elif dire == "Celestial Object Model":
        # Object
        element = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element((By.XPATH, "//select[@name='sel_object']"),other))
        time.sleep(1)
        model_other = Select(browser.find_element_by_name('sel_object'))
        model_other.select_by_visible_text(other)
        browser.execute_script("object_change()")

        # Model
        element = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element((By.XPATH, "//select[@name='sel_model']"),modelo))
        time.sleep(1)
        model_file = Select(browser.find_element_by_name('sel_model'))
        model_file.select_by_visible_text(modelo)
        browser.execute_script("model_change()")

    else:    
        # Model File
        element = WebDriverWait(browser, 10).until(EC.text_to_be_present_in_element((By.XPATH, "//select[@name='sel_model']"),modelo))
        time.sleep(1)
        model_file = Select(browser.find_element_by_name('sel_model'))
        model_file.select_by_visible_text(modelo)
        browser.execute_script("model_change()")

    # Functional
    functional = Select(browser.find_element_by_name('sel_func'))
    functional.select_by_visible_text(func)

    #step
    grid_step = browser.find_element_by_name('grid_step')
    grid_step.clear()
    grid_step.send_keys('0.00000001')

    # height
    ##########################################################################################################################################
    # IMPORTANT NOTE: SINCE mid-July/2018, due to some information provided by *me*, ICGEM has changed the calculation of some functionals;  #
    # Thus, for 'height_anomaly', 'height_anomaly_ell', 'geoid', 'gravity_disturbance', 'gravity_anomaly', 'gravity_anomaly_cl',             #
    #           'gravity_anomaly_bg', 'gravity_earth' and 'water_column' functionals, it is not necessary/allowed to inform "height";        #
    # For these calculations, I advise inserting "0" to the 4th column of the input file.                                                    #
    ##########################################################################################################################################
    if func not in ['height_anomaly', 'height_anomaly_ell', 'geoid', 'gravity_disturbance', 'gravity_anomaly', 'gravity_anomaly_cl', 'gravity_anomaly_bg', 'gravity_earth', 'water_column']:
        height = browser.find_element_by_name('grid_height')
        height.clear()
        height.send_keys(str(hovell))

    # Tide System
    tide = Select(browser.find_element_by_name('sel_tidesys'))
    tide.select_by_visible_text(mare)

    # Zero Degree Term
    zerodeg = browser.find_element_by_name('zero_term')
    if gzero == 'No':
        zerodeg.click()

    # Reference System
    ref = Select(browser.find_element_by_name('sel_refsys'))
    ref.select_by_visible_text(sisref)
    if sisref == '- user defined -':
        radius = Select(browser.find_element_by_name('refsys_radius'))
        radius.clear()
        radius.send_keys(raio)
        gm = Select(browser.find_element_by_name('refsys_gm'))
        gm.clear()
        gm.send_keys(const_gm)
        flat = Select(browser.find_element_by_name('refsys_flat'))
        flat.clear()
        flat.send_keys(achat)
        omega = Select(browser.find_element_by_name('refsys_omega'))
        omega.clear()
        omega.send_keys(achat)

    # Maximal Degree
    if graumax != 9999 and graumax != 0 and graumax != "":
        max_deg = browser.find_element_by_name('trunc_max')
        max_deg.clear()
        max_deg.send_keys(graumax)
        browser.execute_script("trunc_change('max')")
        gc = browser.find_element_by_name('trunc_gentle')
        gc.clear()
        gc.send_keys(graumax)
        browser.execute_script("trunc_change('gentle')")
        #time.sleep(1)

    # Gentle Cut
    if gentlecut != 9999 and gentlecut != 0 and gentlecut != "" and gentlecut != graumax:
        gc = browser.find_element_by_name('trunc_gentle')
        gc.clear()
        gc.send_keys(gentlecut)
        browser.execute_script("trunc_change('gentle')")
        #time.sleep(1)


    # Long
    long_left = browser.find_element_by_name('grid_left')
    long_left.send_keys(Keys.CONTROL + "a");
    long_left.send_keys(Keys.DELETE);
    browser.execute_script("grid_change()")
    long_left.send_keys(str(long))
    long_right = browser.find_element_by_name('grid_right')
    long_right.send_keys(Keys.CONTROL + "a");
    long_right.send_keys(Keys.DELETE);
    browser.execute_script("grid_change()")
    long_right.send_keys(str(long))

    #Lat
    lat_top = browser.find_element_by_name('grid_top')
    lat_top.send_keys(Keys.CONTROL + "a");
    lat_top.send_keys(Keys.DELETE);
    browser.execute_script("grid_change()")
    lat_top.send_keys(str(lat))
    lat_bottom = browser.find_element_by_name('grid_bottom')
    lat_bottom.send_keys(Keys.CONTROL + "a");
    lat_bottom.send_keys(Keys.DELETE);
    browser.execute_script("grid_change()")
    lat_bottom.send_keys(str(lat))
    
    # Start computation
    start = browser.find_element_by_xpath('//input[@type="submit"]')
    start.click()

    WebDriverWait(browser, 10).until(lambda d: len(d.window_handles) == 2)
    browser.switch_to_window(browser.window_handles[1])
    try:
        element_present = EC.text_to_be_present_in_element((By.ID, 'progress_text'),'Done.')
        WebDriverWait(browser, 30).until(element_present)
        nome = '*'+browser.find_element_by_link_text('Download Grid').get_attribute('href').split("http://icgem.gfz-potsdam.de/calcgdf/",1)[1]+'.gdf'
        browser.find_element_by_link_text('Download Grid').click()
        WebDriverWait(browser, 10).until(lambda d: len(d.window_handles) == 3)
        browser.switch_to_window(browser.window_handles[2])
        os.chdir(currfolder)
        while not glob.glob(nome):
            time.sleep(1)
        arq = open(glob.glob(nome)[0], 'r').read()
        os.remove(currfolder+'/'+glob.glob(nome)[0])
        browser.quit()
        display.stop()
        return float(arq[-20:].strip())
    except TimeoutException:
        print("Internet issues. Please, try again.")
        return False

def preenche_series():
    Form2.move(350,150)
    Form2.show()
    display = Display(visible=0, size=(800, 600))
    display.start()
    lista1 = []
    lista2 = []
    currfolder = os.getcwd()
    chromeoptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : currfolder, "extensions_to_open": "gdf","profile.managed_default_content_settings.images":2}
    chromeoptions.add_experimental_option("prefs",prefs)
    browser = webdriver.Chrome(chrome_options=chromeoptions)
    browser.get('http://icgem.gfz-potsdam.de/calc')    

    campo1 = browser.find_element_by_name("sel_series")
    options1 = [x for x in campo1.find_elements_by_tag_name("option")]
    for element in options1:
        a = str(element.get_attribute("value"))
        b = str(element.get_attribute("title"))
        lista1.append(a)
        lista2.append(b)
    z = dict(zip(lista2, lista1))
    z = collections.OrderedDict(z)
    browser.quit()
    display.stop()
    Form2.hide()
    return z


def preenche_modelos():
    global series
    global contador
    ui.modelo.clear()
    ui.func.clear()
    ui.ref_sys.clear()
    if ui.dir.currentItem():
        nome = tipos[ui.dir.currentItem().text()]
        for i in range(len(funcionais[nome])):
            ui.func.addItem(funcionais[nome][i])
        if nome == 'celestial':
            ui.ref_sys.clear()
            ui.radius.clear()
            ui.gm.clear()
            ui.flat.clear()
            ui.omega.clear()
            ui.modelo.addItem('Mars')
            ui.modelo.addItem('Moon (of the Earth)')
            ui.modelo.addItem('Venus')
            ui.other.setEnabled(True)
            ui.tide_sys.setCurrentIndex(0)
            ui.tide_sys.setEnabled(False)
        elif nome == 'series':
            if contador == 0:
                series = preenche_series()
            ui.ref_sys.addItems(refsys_set['Earth'])
            ui.ref_sys.setCurrentIndex(1)
            for key,valor in series.items():
                ui.modelo.addItem(key)
        elif nome == 'topo':
            ui.tide_sys.setCurrentIndex(0)
            ui.tide_sys.setEnabled(False)
            ui.ref_sys.addItems(refsys_set['Earth'])
            ui.ref_sys.setCurrentIndex(1)
            with urllib.request.urlopen("http://icgem.gfz-potsdam.de/json_modelmap/name?modeltype="+nome) as url:
                data = json.loads(url.read().decode())
                for i in range(len(data)):
                    ui.modelo.addItem(data[i][1])    
        else:
            ui.tide_sys.setCurrentIndex(0)
            ui.tide_sys.setEnabled(True)
            ui.ref_sys.addItems(refsys_set['Earth'])
            ui.ref_sys.setCurrentIndex(1)
            with urllib.request.urlopen("http://icgem.gfz-potsdam.de/json_modelmap/name?modeltype="+nome) as url:
                data = json.loads(url.read().decode())
                for i in range(len(data)):
                    ui.modelo.addItem(data[i][1])    

def preenche_other():
    global series
    ui.other.clear()
    if ui.dir.currentItem():
        if ui.dir.currentItem().text() == "Celestial Object Model":
            ui.ref_sys.clear()
            ui.radius.clear()
            ui.gm.clear()
            ui.flat.clear()
            ui.omega.clear()
            if len(ui.modelo.selectedItems()) != 0:
                nome = ui.modelo.currentItem().text()
                if nome != "Mars" and nome != "Venus":
                    x = "http://icgem.gfz-potsdam.de/json_modelmap/name?modeltype=celestial&object=Moon%20%28of%20the%20Earth%29"
                    ui.ref_sys.addItems(refsys_set['Moon'])
                else:
                    x = "http://icgem.gfz-potsdam.de/json_modelmap/name?modeltype=celestial&object="+nome
                    ui.ref_sys.addItems(refsys_set[nome])
                with urllib.request.urlopen(x) as url:
                    data = json.loads(url.read().decode())
                    for i in range(len(data)):
                        ui.other.addItem(data[i][1])
                    ui.other.setEnabled(True)
        elif ui.dir.currentItem().text() == "Model from Series":
            if len(ui.modelo.selectedItems()) != 0:
                nome = ui.modelo.currentItem().text()
                val = series[nome]
                with urllib.request.urlopen("http://icgem.gfz-potsdam.de/json_modelmap/name?modeltype=seriesmodel&series="+val) as url:
                    data = json.loads(url.read().decode())
                    for i in range(len(data)):
                        ui.other.addItem(data[i][1])
                    ui.other.setEnabled(True)
        preenche_spec()    
    else:
        pass
    
def preenche_spec():
    ref=ui.ref_sys.currentText()
    if ref:
        ui.radius.setEnabled(False)
        ui.gm.setEnabled(False)
        ui.flat.setEnabled(False)
        ui.omega.setEnabled(False)
        if ref != "- user defined -":
            ui.radius.setText(ref_sys[ref][0])
            ui.gm.setText(ref_sys[ref][2])
            ui.flat.setText(ref_sys[ref][1])
            ui.omega.setText(ref_sys[ref][3])    
        else:
            ui.radius.setText("")
            ui.gm.setText("")
            ui.flat.setText("")
            ui.omega.setText("")
            ui.radius.setEnabled(True)
            ui.gm.setEnabled(True)
            ui.flat.setEnabled(True)
            ui.omega.setEnabled(True)
    
    
def gera_modelos(fileName, dire, other, modelo, func, mare, sisref, raio, const_gm, achat, velrot, gzero, gentlecut, graumax):
    global SAIDA
    i = 0
    pontos = le_xls(fileName,0)
    numrows = len(pontos)
    print ("\nStarting calculation...")
    printProgressBar(i, numrows, prefix = 'Progress:', suffix = 'Complete', length = 50)
    d = conv_tempo(20*numrows)
    ui.label_tempoestimado.setText(str(d))
    ui.label_pontos.setText("0 of "+str(len(pontos)))
    geo = [[0 for x in range(2)] for y in range(numrows)]
    Form2.move(350,150)
    Form2.show()
    inicio = time.time()
    while i < numrows:
        ui.progresso.setValue(i*100/len(pontos))
        ui.label_pontos.setText(str(i+1)+" of "+str(len(pontos)))
        geo[i][0] = pontos[i][0]
        geo[i][1] = gera_grid_ponto(dire, other, modelo, func, mare, pontos[i][1],pontos[i][2],pontos[i][3], sisref, raio, const_gm, achat, velrot, gzero, gentlecut, graumax)
        escreve_lista_arq('temp.txt',geo[i])
        d = conv_tempo(int(float(time.time())-(float(inicio))))
        ui.label_tempotranscorrido.setText(str(d))
        print(geo[i])
        i+=1
        printProgressBar(i, numrows, prefix = 'Progress:', suffix = 'Complete', length = 50)
    print('Extract generated on '+ gera_nome_saida(fileName) +' for GGM '+modelo+', Functional '+func+' and Max Degree '+str(graumax)+'.\n')
    Form2.hide()
    grava_xls(gera_nome_saida(fileName),geo,1,2)
    ui.progresso.setValue(100)
    alerta("Calculation finished!", "Click the 'Open results' button to view the results.", 1)
    SAIDA = "ok"

    
def verifica_campos():
    global fileName
    verifica = True
    
    # verifica arquivo
    if fileName == '':
        verifica = False
        
    # verifica dir
    if len(ui.dir.selectedItems()) != 0:
        dire = ui.dir.currentItem().text()
    else:
        dire = ""
        verifica = False
    
    #verifica other
    if (dire == "Model from Series" or dire == "Celestial Object Model") and len(ui.other.selectedItems()) != 0:
        other = ui.other.currentItem().text()
    elif (dire != "Model from Series" and dire != "Celestial Object Model"):
        other = ""
    else:
        verifica = False
    
    # verifica modelo
    if len(ui.modelo.selectedItems()) != 0:
        modelo = ui.modelo.currentItem().text()
    else:
        verifica = False
    
    # verifica func
    if len(ui.func.selectedItems()) != 0:
        func = ui.func.currentItem().text()
    else:
        verifica = False
        
    # verifica mare
    mare = ui.tide_sys.currentText()
    
    # verifica sisref
    sisref = ui.ref_sys.currentText()
    raio = ui.radius.text()
    const_gm = ui.gm.text()
    achat = ui.flat.text()
    velrot = ui.omega.text()
    
    # verifica grau zero
    gzero = ui.zero_deg.currentText()
    
    # verifica gentlecut e graumax
    gentlecut = ui.cut.value()
    graumax = ui.grau.value()
    
    if gentlecut > graumax and gentlecut!=9999:
        verifica = False
    
    if verifica:
        gera_modelos(fileName, dire, other, modelo, func, mare, sisref, raio, const_gm, achat, velrot, gzero, gentlecut, graumax)
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
    

# main
if __name__ == "__main__":
    global contador
    contador = 0
    global fileName
    global SAIDA
    global series
    fileName = ''
    SAIDA = ''
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form2 = QtWidgets.QWidget()
    ui2 = Ui_Loading()
    ui2.setupUi(Form2)
    if(testa_con()):
        Form.move(300,100)
        Form.show()
        for value,key in tipos.items():
	        ui.dir.addItem(value)
        ui.zero_deg.addItems(['Yes','No'])
        ui.tide_sys.addItems(['use model\'s system','tide free','zero tide','mean tide'])
        ui.ref_sys.currentIndexChanged.connect(preenche_spec)
        ui.dir.currentItemChanged.connect(preenche_modelos)
        ui.modelo.itemSelectionChanged.connect(preenche_other)
        ui.pushButton_2.clicked.connect(testa_con)
        ui.fechar.clicked.connect(fecha)
        ui.limpar.clicked.connect(limpa)
        ui.abrir.clicked.connect(openFileNameDialog)
        ui.gerar.clicked.connect(verifica_campos)
        ui.abrir_result.clicked.connect(abre_resultado)
    else:
        fecha()
    sys.exit(app.exec_())
