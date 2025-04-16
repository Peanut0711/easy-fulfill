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
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QMainWindow, QMenu,
    QMenuBar, QPlainTextEdit, QPushButton, QSizePolicy,
    QSpacerItem, QStatusBar, QTabWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(732, 682)
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
        self.tabWidget = QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName(u"tabWidget")
        self.tab = QWidget()
        self.tab.setObjectName(u"tab")
        self.filePathLabel = QLabel(self.tab)
        self.filePathLabel.setObjectName(u"filePathLabel")
        self.filePathLabel.setGeometry(QRect(190, 60, 491, 40))
        self.filePathLabel.setMinimumSize(QSize(300, 40))
        self.filePathLabel.setMaximumSize(QSize(500, 40))
        font = QFont()
        font.setFamilies([u"\ub9d1\uc740 \uace0\ub515"])
        font.setPointSize(10)
        font.setBold(True)
        self.filePathLabel.setFont(font)
        self.filePathLabel.setFrameShape(QFrame.Shape.Panel)
        self.filePathLabel.setFrameShadow(QFrame.Shadow.Sunken)
        self.filePathLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.filePathLabel.setMargin(5)
        self.plainTextEdit = QPlainTextEdit(self.tab)
        self.plainTextEdit.setObjectName(u"plainTextEdit")
        self.plainTextEdit.setGeometry(QRect(20, 130, 661, 431))
        self.label_logo = QLabel(self.tab)
        self.label_logo.setObjectName(u"label_logo")
        self.label_logo.setGeometry(QRect(30, 60, 151, 41))
        font1 = QFont()
        font1.setPointSize(10)
        font1.setBold(True)
        self.label_logo.setFont(font1)
        self.label_logo.setFrameShape(QFrame.Shape.StyledPanel)
        self.label_logo.setFrameShadow(QFrame.Shadow.Plain)
        self.label_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QWidget()
        self.tab_2.setObjectName(u"tab_2")
        self.verticalLayout_2 = QVBoxLayout(self.tab_2)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.productFileLabel = QLabel(self.tab_2)
        self.productFileLabel.setObjectName(u"productFileLabel")
        self.productFileLabel.setMinimumSize(QSize(300, 40))
        self.productFileLabel.setMaximumSize(QSize(500, 40))
        self.productFileLabel.setFont(font)
        self.productFileLabel.setFrameShape(QFrame.Shape.Panel)
        self.productFileLabel.setFrameShadow(QFrame.Shadow.Sunken)
        self.productFileLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.productFileLabel.setMargin(5)

        self.horizontalLayout_3.addWidget(self.productFileLabel)

        self.selectProductFileButton = QPushButton(self.tab_2)
        self.selectProductFileButton.setObjectName(u"selectProductFileButton")
        self.selectProductFileButton.setMinimumSize(QSize(120, 40))
        self.selectProductFileButton.setMaximumSize(QSize(120, 40))
        self.selectProductFileButton.setFont(font1)

        self.horizontalLayout_3.addWidget(self.selectProductFileButton)


        self.verticalLayout_2.addLayout(self.horizontalLayout_3)

        self.categoryTableWidget = QTableWidget(self.tab_2)
        if (self.categoryTableWidget.columnCount() < 6):
            self.categoryTableWidget.setColumnCount(6)
        __qtablewidgetitem = QTableWidgetItem()
        self.categoryTableWidget.setHorizontalHeaderItem(0, __qtablewidgetitem)
        __qtablewidgetitem1 = QTableWidgetItem()
        self.categoryTableWidget.setHorizontalHeaderItem(1, __qtablewidgetitem1)
        __qtablewidgetitem2 = QTableWidgetItem()
        self.categoryTableWidget.setHorizontalHeaderItem(2, __qtablewidgetitem2)
        __qtablewidgetitem3 = QTableWidgetItem()
        self.categoryTableWidget.setHorizontalHeaderItem(3, __qtablewidgetitem3)
        __qtablewidgetitem4 = QTableWidgetItem()
        self.categoryTableWidget.setHorizontalHeaderItem(4, __qtablewidgetitem4)
        __qtablewidgetitem5 = QTableWidgetItem()
        self.categoryTableWidget.setHorizontalHeaderItem(5, __qtablewidgetitem5)
        self.categoryTableWidget.setObjectName(u"categoryTableWidget")
        font2 = QFont()
        font2.setFamilies([u"\ub9d1\uc740 \uace0\ub515"])
        font2.setPointSize(9)
        self.categoryTableWidget.setFont(font2)
        self.categoryTableWidget.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.categoryTableWidget.setAlternatingRowColors(True)
        self.categoryTableWidget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.categoryTableWidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.categoryTableWidget.setSortingEnabled(True)
        self.categoryTableWidget.horizontalHeader().setStretchLastSection(True)

        self.verticalLayout_2.addWidget(self.categoryTableWidget)

        self.horizontalLayout_4 = QHBoxLayout()
        self.horizontalLayout_4.setObjectName(u"horizontalLayout_4")
        self.horizontalSpacer_5 = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.horizontalLayout_4.addItem(self.horizontalSpacer_5)

        self.categorizeButton = QPushButton(self.tab_2)
        self.categorizeButton.setObjectName(u"categorizeButton")
        self.categorizeButton.setMinimumSize(QSize(150, 40))
        self.categorizeButton.setMaximumSize(QSize(150, 40))
        self.categorizeButton.setFont(font1)

        self.horizontalLayout_4.addWidget(self.categorizeButton)

        self.exportCategoryButton = QPushButton(self.tab_2)
        self.exportCategoryButton.setObjectName(u"exportCategoryButton")
        self.exportCategoryButton.setMinimumSize(QSize(150, 40))
        self.exportCategoryButton.setMaximumSize(QSize(150, 40))
        self.exportCategoryButton.setFont(font1)

        self.horizontalLayout_4.addWidget(self.exportCategoryButton)

        self.horizontalSpacer_6 = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        self.horizontalLayout_4.addItem(self.horizontalSpacer_6)


        self.verticalLayout_2.addLayout(self.horizontalLayout_4)

        self.tabWidget.addTab(self.tab_2, "")

        self.verticalLayout.addWidget(self.tabWidget)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 732, 33))
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

        self.tabWidget.setCurrentIndex(0)


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
        self.label_logo.setText(QCoreApplication.translate("MainWindow", u"\uc8fc\ubb38 \uc815\ubcf4 \uc5c6\uc74c", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), QCoreApplication.translate("MainWindow", u"\uc8fc\ubb38\ucc98\ub9ac", None))
        self.productFileLabel.setText(QCoreApplication.translate("MainWindow", u"\uc120\ud0dd\ub41c \ud30c\uc77c \uc5c6\uc74c", None))
        self.selectProductFileButton.setText(QCoreApplication.translate("MainWindow", u"\uc0c1\ud488 \ud30c\uc77c \uc5f4\uae30", None))
        ___qtablewidgetitem = self.categoryTableWidget.horizontalHeaderItem(0)
        ___qtablewidgetitem.setText(QCoreApplication.translate("MainWindow", u"\uce74\ud14c\uace0\ub9ac", None));
        ___qtablewidgetitem1 = self.categoryTableWidget.horizontalHeaderItem(1)
        ___qtablewidgetitem1.setText(QCoreApplication.translate("MainWindow", u"\uc911\ubd84\ub958", None));
        ___qtablewidgetitem2 = self.categoryTableWidget.horizontalHeaderItem(2)
        ___qtablewidgetitem2.setText(QCoreApplication.translate("MainWindow", u"\uc18c\ubd84\ub958", None));
        ___qtablewidgetitem3 = self.categoryTableWidget.horizontalHeaderItem(3)
        ___qtablewidgetitem3.setText(QCoreApplication.translate("MainWindow", u"\uc0c1\ud488\uba85", None));
        ___qtablewidgetitem4 = self.categoryTableWidget.horizontalHeaderItem(4)
        ___qtablewidgetitem4.setText(QCoreApplication.translate("MainWindow", u"\uc774\ubbf8\uc9c0", None));
        ___qtablewidgetitem5 = self.categoryTableWidget.horizontalHeaderItem(5)
        ___qtablewidgetitem5.setText(QCoreApplication.translate("MainWindow", u"\ube44\uace0", None));
        self.categorizeButton.setText(QCoreApplication.translate("MainWindow", u"\uc0c1\ud488 \ubd84\ub958", None))
        self.exportCategoryButton.setText(QCoreApplication.translate("MainWindow", u"\uc5d1\uc140\ub85c \ub0b4\ubcf4\ub0b4\uae30", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), QCoreApplication.translate("MainWindow", u"\uc0c1\ud488\ubd84\ub958", None))
        self.menuFile.setTitle(QCoreApplication.translate("MainWindow", u"\ud30c\uc77c(&F)", None))
        self.menuHelp.setTitle(QCoreApplication.translate("MainWindow", u"\ub3c4\uc6c0\ub9d0(&H)", None))
        self.statusbar.setStyleSheet(QCoreApplication.translate("MainWindow", u"QStatusBar {\n"
"    border-top: 1px solid #CCCCCC;\n"
"}", None))
    # retranslateUi

