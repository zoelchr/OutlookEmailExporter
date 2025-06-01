# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'ui_main_window.ui'
##
## Created by: Qt User Interface Compiler version 6.9.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QAction, QBrush, QColor, QConicalGradient,
    QCursor, QFont, QFontDatabase, QGradient,
    QIcon, QImage, QKeySequence, QLinearGradient,
    QPainter, QPalette, QPixmap, QRadialGradient,
    QTransform)
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QGridLayout,
    QHeaderView, QLabel, QLayout, QMainWindow,
    QMenu, QMenuBar, QPushButton, QSizePolicy,
    QSpacerItem, QStatusBar, QTableView, QVBoxLayout,
    QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(1000, 603)
        MainWindow.setMinimumSize(QSize(1000, 600))
        self.actionExit = QAction(MainWindow)
        self.actionExit.setObjectName(u"actionExit")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.gridLayout = QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName(u"gridLayout")
        self.pushButton_Export_PDF = QPushButton(self.centralwidget)
        self.pushButton_Export_PDF.setObjectName(u"pushButton_Export_PDF")
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_Export_PDF.sizePolicy().hasHeightForWidth())
        self.pushButton_Export_PDF.setSizePolicy(sizePolicy)
        self.pushButton_Export_PDF.setMinimumSize(QSize(120, 0))
        self.pushButton_Export_PDF.setMaximumSize(QSize(100, 23))

        self.gridLayout.addWidget(self.pushButton_Export_PDF, 2, 1, 1, 1)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setSizeConstraint(QLayout.SizeConstraint.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(5, 5, 5, 5)
        self.label_Postfach = QLabel(self.centralwidget)
        self.label_Postfach.setObjectName(u"label_Postfach")

        self.verticalLayout.addWidget(self.label_Postfach)

        self.comboBox_Postfach = QComboBox(self.centralwidget)
        self.comboBox_Postfach.setObjectName(u"comboBox_Postfach")
        self.comboBox_Postfach.setMinimumSize(QSize(400, 23))

        self.verticalLayout.addWidget(self.comboBox_Postfach)

        self.label_Verzeichnis = QLabel(self.centralwidget)
        self.label_Verzeichnis.setObjectName(u"label_Verzeichnis")

        self.verticalLayout.addWidget(self.label_Verzeichnis)

        self.comboBox_Verzeichnis = QComboBox(self.centralwidget)
        self.comboBox_Verzeichnis.setObjectName(u"comboBox_Verzeichnis")
        self.comboBox_Verzeichnis.setMinimumSize(QSize(100, 23))

        self.verticalLayout.addWidget(self.comboBox_Verzeichnis)


        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 5)

        self.pushButton_Export_MSG = QPushButton(self.centralwidget)
        self.pushButton_Export_MSG.setObjectName(u"pushButton_Export_MSG")
        sizePolicy.setHeightForWidth(self.pushButton_Export_MSG.sizePolicy().hasHeightForWidth())
        self.pushButton_Export_MSG.setSizePolicy(sizePolicy)
        self.pushButton_Export_MSG.setMinimumSize(QSize(120, 0))
        self.pushButton_Export_MSG.setMaximumSize(QSize(100, 23))

        self.gridLayout.addWidget(self.pushButton_Export_MSG, 2, 0, 1, 1)

        self.verticalLayout_2 = QVBoxLayout()
        self.verticalLayout_2.setSpacing(5)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.verticalLayout_2.setSizeConstraint(QLayout.SizeConstraint.SetDefaultConstraint)
        self.verticalLayout_2.setContentsMargins(5, 5, 5, 5)
        self.label_Tabelle = QLabel(self.centralwidget)
        self.label_Tabelle.setObjectName(u"label_Tabelle")

        self.verticalLayout_2.addWidget(self.label_Tabelle)

        self.tableView_Emails = QTableView(self.centralwidget)
        self.tableView_Emails.setObjectName(u"tableView_Emails")

        self.verticalLayout_2.addWidget(self.tableView_Emails)

        self.label_Zielverzeichnis = QLabel(self.centralwidget)
        self.label_Zielverzeichnis.setObjectName(u"label_Zielverzeichnis")

        self.verticalLayout_2.addWidget(self.label_Zielverzeichnis)

        self.comboBox_Exportziel = QComboBox(self.centralwidget)
        self.comboBox_Exportziel.setObjectName(u"comboBox_Exportziel")

        self.verticalLayout_2.addWidget(self.comboBox_Exportziel)


        self.gridLayout.addLayout(self.verticalLayout_2, 1, 0, 1, 9)

        self.checkBox_Change_Filename = QCheckBox(self.centralwidget)
        self.checkBox_Change_Filename.setObjectName(u"checkBox_Change_Filename")

        self.gridLayout.addWidget(self.checkBox_Change_Filename, 2, 5, 1, 1)

        self.horizontalSpacer = QSpacerItem(20, 20, QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Minimum)

        self.gridLayout.addItem(self.horizontalSpacer, 2, 3, 1, 1)

        self.checkBox_Change_Filedate = QCheckBox(self.centralwidget)
        self.checkBox_Change_Filedate.setObjectName(u"checkBox_Change_Filedate")

        self.gridLayout.addWidget(self.checkBox_Change_Filedate, 2, 4, 1, 1)

        self.horizontalSpacer_2 = QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.gridLayout.addItem(self.horizontalSpacer_2, 2, 7, 1, 1)

        self.pushButton_Export_Both = QPushButton(self.centralwidget)
        self.pushButton_Export_Both.setObjectName(u"pushButton_Export_Both")
        sizePolicy.setHeightForWidth(self.pushButton_Export_Both.sizePolicy().hasHeightForWidth())
        self.pushButton_Export_Both.setSizePolicy(sizePolicy)
        self.pushButton_Export_Both.setMinimumSize(QSize(120, 0))
        self.pushButton_Export_Both.setMaximumSize(QSize(150, 23))

        self.gridLayout.addWidget(self.pushButton_Export_Both, 2, 2, 1, 1)

        self.pushButton_Exit = QPushButton(self.centralwidget)
        self.pushButton_Exit.setObjectName(u"pushButton_Exit")
        self.pushButton_Exit.setEnabled(True)
        self.pushButton_Exit.setMaximumSize(QSize(100, 23))

        self.gridLayout.addWidget(self.pushButton_Exit, 2, 8, 1, 1)

        self.checkBox_Overwrite_File = QCheckBox(self.centralwidget)
        self.checkBox_Overwrite_File.setObjectName(u"checkBox_Overwrite_File")

        self.gridLayout.addWidget(self.checkBox_Overwrite_File, 2, 6, 1, 1)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 1000, 33))
        self.menuExit = QMenu(self.menubar)
        self.menuExit.setObjectName(u"menuExit")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.menubar.addAction(self.menuExit.menuAction())
        self.menuExit.addAction(self.actionExit)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.actionExit.setText(QCoreApplication.translate("MainWindow", u"Beenden", None))
        self.pushButton_Export_PDF.setText(QCoreApplication.translate("MainWindow", u"Export PDF", None))
        self.label_Postfach.setText(QCoreApplication.translate("MainWindow", u"Postfach", None))
        self.label_Verzeichnis.setText(QCoreApplication.translate("MainWindow", u"Verzeichnis", None))
        self.pushButton_Export_MSG.setText(QCoreApplication.translate("MainWindow", u"Export MSG", None))
        self.label_Tabelle.setText(QCoreApplication.translate("MainWindow", u"Emails", None))
        self.label_Zielverzeichnis.setText(QCoreApplication.translate("MainWindow", u"Zielverzeichnis", None))
        self.checkBox_Change_Filename.setText(QCoreApplication.translate("MainWindow", u"Dateiname \u00e4ndern?", None))
        self.checkBox_Change_Filedate.setText(QCoreApplication.translate("MainWindow", u"Dateidatum \u00e4ndern?", None))
        self.pushButton_Export_Both.setText(QCoreApplication.translate("MainWindow", u"Export MSG && PDF", None))
        self.pushButton_Exit.setText(QCoreApplication.translate("MainWindow", u"Beenden", None))
        self.checkBox_Overwrite_File.setText(QCoreApplication.translate("MainWindow", u"Dateien \u00fcberschreiben?", None))
        self.menuExit.setTitle(QCoreApplication.translate("MainWindow", u"File", None))
    # retranslateUi

