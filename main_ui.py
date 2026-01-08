# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main.ui'
##
## Created by: Qt User Interface Compiler version 6.10.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QGroupBox,
    QHBoxLayout, QLineEdit, QListView, QMainWindow,
    QMenuBar, QPushButton, QSizePolicy, QStatusBar,
    QTextBrowser, QTextEdit, QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(656, 497)
        MainWindow.setMinimumSize(QSize(532, 461))
        icon = QIcon()
        icon.addFile(u"logo.png", QSize(), QIcon.Mode.Normal, QIcon.State.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet(u"* {\n"
"font-family: \"Noto Sans\";\n"
"font-size: 9pt;\n"
"}\n"
"\n"
"QMainWindow > QWidget {\n"
"    background-color: #f0f0f0;\n"
"}\n"
"\n"
"QListView {\n"
"background-color: #ffffff;\n"
"}\n"
"\n"
"/* RadioButton \uae30\ubcf8 indicator \uc2a4\ud0c0\uc77c */\n"
"QRadioButton::indicator:unchecked {\n"
"background-color: white;\n"
"border: 2px solid white;\n"
"}\n"
"/* \uc120\ud0dd \uc2dc \uc0c9\uae54 */\n"
"QRadioButton::indicator:checked {\n"
"background-color: black;\n"
"border: 2px solid white;\n"
"}")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.grpSavePath = QGroupBox(self.centralwidget)
        self.grpSavePath.setObjectName(u"grpSavePath")
        self.grpSavePath.setGeometry(QRect(20, 190, 440, 100))
        font = QFont()
        font.setFamilies([u"Noto Sans"])
        font.setPointSize(9)
        font.setBold(False)
        self.grpSavePath.setFont(font)
        self.lineEditSavePath = QLineEdit(self.grpSavePath)
        self.lineEditSavePath.setObjectName(u"lineEditSavePath")
        self.lineEditSavePath.setGeometry(QRect(10, 30, 420, 26))
        self.lineEditSavePath.setFont(font)
        self.widget = QWidget(self.grpSavePath)
        self.widget.setObjectName(u"widget")
        self.widget.setGeometry(QRect(10, 60, 420, 30))
        self.horizontalLayout = QHBoxLayout(self.widget)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.btnBrowsePath = QPushButton(self.widget)
        self.btnBrowsePath.setObjectName(u"btnBrowsePath")
        self.btnBrowsePath.setFont(font)

        self.horizontalLayout.addWidget(self.btnBrowsePath)

        self.btnOpenPath = QPushButton(self.widget)
        self.btnOpenPath.setObjectName(u"btnOpenPath")
        self.btnOpenPath.setFont(font)

        self.horizontalLayout.addWidget(self.btnOpenPath)

        self.btnStart = QPushButton(self.widget)
        self.btnStart.setObjectName(u"btnStart")
        self.btnStart.setFont(font)

        self.horizontalLayout.addWidget(self.btnStart)

        self.txtLogOutput = QTextBrowser(self.centralwidget)
        self.txtLogOutput.setObjectName(u"txtLogOutput")
        self.txtLogOutput.setGeometry(QRect(20, 300, 600, 150))
        self.txtLogOutput.setFont(font)
        self.groupBoxfiles = QGroupBox(self.centralwidget)
        self.groupBoxfiles.setObjectName(u"groupBoxfiles")
        self.groupBoxfiles.setGeometry(QRect(20, 10, 600, 170))
        self.listViewafter = QListView(self.groupBoxfiles)
        self.listViewafter.setObjectName(u"listViewafter")
        self.listViewafter.setGeometry(QRect(310, 30, 280, 130))
        self.listViewafter.setDragEnabled(True)
        self.listViewafter.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.listViewafter.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.listViewbefore = QListView(self.groupBoxfiles)
        self.listViewbefore.setObjectName(u"listViewbefore")
        self.listViewbefore.setGeometry(QRect(10, 30, 280, 130))
        self.listViewbefore.setDragEnabled(True)
        self.listViewbefore.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.listViewbefore.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setObjectName(u"groupBox")
        self.groupBox.setGeometry(QRect(470, 190, 150, 100))
        self.checkBoxExcel = QCheckBox(self.groupBox)
        self.checkBoxExcel.setObjectName(u"checkBoxExcel")
        self.checkBoxExcel.setGeometry(QRect(10, 70, 111, 23))
        self.textEditauthor = QTextEdit(self.groupBox)
        self.textEditauthor.setObjectName(u"textEditauthor")
        self.textEditauthor.setGeometry(QRect(10, 30, 120, 30))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 656, 23))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.grpSavePath.setTitle(QCoreApplication.translate("MainWindow", u"\uc800\uc7a5 \uc704\uce58", None))
        self.btnBrowsePath.setText(QCoreApplication.translate("MainWindow", u"\uacbd\ub85c \uc9c0\uc815", None))
#if QT_CONFIG(shortcut)
        self.btnBrowsePath.setShortcut(QCoreApplication.translate("MainWindow", u"Ctrl+S", None))
#endif // QT_CONFIG(shortcut)
        self.btnOpenPath.setText(QCoreApplication.translate("MainWindow", u"\uacbd\ub85c \uc5f4\uae30", None))
#if QT_CONFIG(shortcut)
        self.btnOpenPath.setShortcut(QCoreApplication.translate("MainWindow", u"Ctrl+O", None))
#endif // QT_CONFIG(shortcut)
        self.btnStart.setText(QCoreApplication.translate("MainWindow", u"\uc791\uc5c5 \uc2dc\uc791(F5)", None))
#if QT_CONFIG(shortcut)
        self.btnStart.setShortcut(QCoreApplication.translate("MainWindow", u"F5", None))
#endif // QT_CONFIG(shortcut)
        self.groupBoxfiles.setTitle(QCoreApplication.translate("MainWindow", u"\ube44\uad50 \ud30c\uc77c(\uc804/\ud6c4)", None))
        self.groupBox.setTitle(QCoreApplication.translate("MainWindow", u"\uae30\ud0c0", None))
        self.checkBoxExcel.setText(QCoreApplication.translate("MainWindow", u"\uc5d1\uc140 \ube44\uad50\ud45c \uc0dd\uc131", None))
#if QT_CONFIG(shortcut)
        self.checkBoxExcel.setShortcut(QCoreApplication.translate("MainWindow", u"E", None))
#endif // QT_CONFIG(shortcut)
        self.textEditauthor.setPlaceholderText(QCoreApplication.translate("MainWindow", u"\uc791\uc5c5\uc790\uba85", None))
    # retranslateUi

