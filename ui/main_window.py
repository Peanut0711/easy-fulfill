# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main_window.ui'
##
## Created by: Qt User Interface Compiler version 6.8.1
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
from PySide6.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel,
    QMainWindow, QMenu, QMenuBar, QPushButton,
    QSizePolicy, QSpacerItem, QStatusBar, QVBoxLayout,
    QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(780, 560)
        self.actionOpenExcel = QAction(MainWindow)
        self.actionOpenExcel.setObjectName(u"actionOpenExcel")
        self.actionExit = QAction(MainWindow)
        self.actionExit.setObjectName(u"actionExit")
        self.actionAbout = QAction(MainWindow)
        self.actionAbout.setObjectName(u"actionAbout")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.verticalLayout = QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.horizontalSpacer_3 = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.horizontalLayout.addItem(self.horizontalSpacer_3)

        self.filePathLabel = QLabel(self.centralwidget)
        self.filePathLabel.setObjectName(u"filePathLabel")
        self.filePathLabel.setMinimumSize(QSize(300, 40))
        self.filePathLabel.setMaximumSize(QSize(500, 40))
        font = QFont()
        font.setFamilies([u"\ub9d1\uc740 \uace0\ub515"])
        font.setPointSize(10)
        font.setBold(True)
        self.filePathLabel.setFont(font)
        self.filePathLabel.setFrameShape(QFrame.Shape.Panel)
        self.filePathLabel.setFrameShadow(QFrame.Shadow.Sunken)
        self.filePathLabel.setMargin(5)

        self.horizontalLayout.addWidget(self.filePathLabel)

        self.selectFileButton = QPushButton(self.centralwidget)
        self.selectFileButton.setObjectName(u"selectFileButton")
        self.selectFileButton.setMinimumSize(QSize(200, 40))
        self.selectFileButton.setMaximumSize(QSize(200, 40))
        font1 = QFont()
        font1.setPointSize(10)
        font1.setBold(True)
        self.selectFileButton.setFont(font1)

        self.horizontalLayout.addWidget(self.selectFileButton)

        self.horizontalSpacer_4 = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.horizontalLayout.addItem(self.horizontalSpacer_4)


        self.verticalLayout.addLayout(self.horizontalLayout)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.horizontalSpacer = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.horizontalLayout_2.addItem(self.horizontalSpacer)

        self.generateButton = QPushButton(self.centralwidget)
        self.generateButton.setObjectName(u"generateButton")
        self.generateButton.setMinimumSize(QSize(200, 40))
        self.generateButton.setMaximumSize(QSize(200, 40))
        self.generateButton.setFont(font1)

        self.horizontalLayout_2.addWidget(self.generateButton)

        self.exportInvoiceButton = QPushButton(self.centralwidget)
        self.exportInvoiceButton.setObjectName(u"exportInvoiceButton")
        self.exportInvoiceButton.setMinimumSize(QSize(200, 40))
        self.exportInvoiceButton.setMaximumSize(QSize(200, 40))
        self.exportInvoiceButton.setFont(font1)

        self.horizontalLayout_2.addWidget(self.exportInvoiceButton)

        self.horizontalSpacer_2 = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.horizontalLayout_2.addItem(self.horizontalSpacer_2)


        self.verticalLayout.addLayout(self.horizontalLayout_2)

        self.verticalSpacer = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

        self.verticalLayout.addItem(self.verticalSpacer)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 780, 33))
        self.menuFile = QMenu(self.menubar)
        self.menuFile.setObjectName(u"menuFile")
        self.menuHelp = QMenu(self.menubar)
        self.menuHelp.setObjectName(u"menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menuFile.addAction(self.actionOpenExcel)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionAbout)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"Easy Fulfill - \uc8fc\ubb38 \ucc98\ub9ac \uc2dc\uc2a4\ud15c", None))
        MainWindow.setStyleSheet(QCoreApplication.translate("MainWindow", u"QMainWindow::separator {\n"
"    height: 1px;\n"
"    background: #CCCCCC;\n"
"    margin: 0px;\n"
"    padding: 0px;\n"
"}", None))
        self.actionOpenExcel.setText(QCoreApplication.translate("MainWindow", u"\uc5d1\uc140 \ud30c\uc77c \uc5f4\uae30(&O)", None))
#if QT_CONFIG(shortcut)
        self.actionOpenExcel.setShortcut(QCoreApplication.translate("MainWindow", u"Ctrl+O", None))
#endif // QT_CONFIG(shortcut)
        self.actionExit.setText(QCoreApplication.translate("MainWindow", u"\uc885\ub8cc(&X)", None))
#if QT_CONFIG(shortcut)
        self.actionExit.setShortcut(QCoreApplication.translate("MainWindow", u"Alt+F4", None))
#endif // QT_CONFIG(shortcut)
        self.actionAbout.setText(QCoreApplication.translate("MainWindow", u"\ud504\ub85c\uadf8\ub7a8 \uc815\ubcf4(&A)", None))
        self.filePathLabel.setText(QCoreApplication.translate("MainWindow", u"\uc120\ud0dd\ub41c \ud30c\uc77c \uc5c6\uc74c", None))
        self.selectFileButton.setText(QCoreApplication.translate("MainWindow", u"\uc5d1\uc140 \ud30c\uc77c \uc120\ud0dd", None))
        self.generateButton.setText(QCoreApplication.translate("MainWindow", u"\uc791\uc5c5\uc9c0\uc2dc\uc11c \uc0dd\uc131", None))
        self.exportInvoiceButton.setText(QCoreApplication.translate("MainWindow", u"\uc1a1\uc7a5 \uc5d1\uc140", None))
        self.menuFile.setTitle(QCoreApplication.translate("MainWindow", u"\ud30c\uc77c(&F)", None))
        self.menuHelp.setTitle(QCoreApplication.translate("MainWindow", u"\ub3c4\uc6c0\ub9d0(&H)", None))
        self.statusbar.setStyleSheet(QCoreApplication.translate("MainWindow", u"QStatusBar {\n"
"    border-top: 1px solid #CCCCCC;\n"
"}", None))
    # retranslateUi

