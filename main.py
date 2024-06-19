import os
from PyQt5 import QtCore, QtGui, QtWidgets
from settings import uiSettings
from ProgressBar import Ui_progresWIndow
from pyautocad import Autocad, APoint

class pilihExcel(QtWidgets.QLabel):
    clicked = QtCore.pyqtSignal()
    def __init__(self):
        super().__init__()
        self.setMinimumSize(QtCore.QSize(100, 200))
        self.setText("")
        self.setPixmap(QtGui.QPixmap("logo excel 20% size.png"))
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setObjectName("pilihExcel")

    def mouseReleaseEvent(self, QMouseEvent):
        self.clicked.emit()

class mainWindow(QtWidgets.QWidget):
    clicked = QtCore.pyqtSignal()
    def __init__(self):
        super().__init__()

        self.path = ""
        self.fileName = ""
        self.setAcceptDrops(True)
        self.setObjectName("Form")
        self.resize(500, 400)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")

        self.setWindowIcon(QtGui.QIcon('main icon.png'))

        self.information = QtWidgets.QLabel(self)
        self.information.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.information.setObjectName("label")
        self.information.setStyleSheet("background-color: rgb(255, 255, 255); color: rgb(0, 80, 0)")
        self.horizontalLayout_2.addWidget(self.information)

        self.setting = QtWidgets.QPushButton(self)
        self.setting.setMaximumSize(QtCore.QSize(30, 30))
        self.setting.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.setting.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("logo settings.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setting.setIcon(icon)
        self.setting.setObjectName("setting")

        self.horizontalLayout_2.addWidget(self.setting)
        self.help = QtWidgets.QPushButton(self)
        self.help.setMaximumSize(QtCore.QSize(30, 30))
        self.help.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.help.setObjectName("help")
        self.horizontalLayout_2.addWidget(self.help)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)

        self.pilihExcel = pilihExcel()
        self.pilihExcel.setStyleSheet("background-color: rgb(215, 255, 215); border: 1px dashed;")
        self.verticalLayout_2.addWidget(self.pilihExcel)

        self.mulai = QtWidgets.QPushButton(self)
        self.mulai.setStyleSheet("background-color: rgb(0, 177, 0);\n"
                                 "color: rgb(255, 255, 255);")
        self.mulai.setObjectName("mulai")
        self.verticalLayout_2.addWidget(self.mulai)

        self.pilihExcel.clicked.connect(self.cariFile)
        self.clicked.connect(self.cariFile)
        self.setting.clicked.connect(self.bukaSetting)
        self.mulai.clicked.connect(self.startAutocad)
        self.help.clicked.connect(self.openHelp)

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Interlocking Table Excel to Autocad"))
        self.help.setText(_translate("Form", "?"))
        self.mulai.setText(_translate("Form", "START"))
        self.information.setText(_translate("Form", " Tarik File Excel Kedalam Kotak Hijau Atau Klik Untuk Mencari File"))

    def cariFile(self):
        fname = QtWidgets.QFileDialog.getOpenFileName(None, 'Open file', os.getcwd(), "Excel Files (*.xlsx *.xlsm *xltx *xltm)")
        if fname[0]:
            self.information.setText(fname[0])
            self.fileName = fname[0].split("/")[-1]
            self.path = fname[0]

    def dragEnterEvent(self, event):
        if event.mimeData().hasImage:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasImage:
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData():
            event.setDropAction(QtCore.Qt.CopyAction)
            file_path = event.mimeData().urls()[0].toLocalFile()
            self.information.setText(file_path)
            self.path = file_path
            self.fileName = self.path.split("/")[-1]
            event.accept()
        else:
            event.ignore()

    def bukaSetting(self):
        self.hide()
        settingDialog = uiSettings()
        settingDialog.setModal(True)
        settingDialog.exec()
        self.show()

    def openHelp(self):
        try:
            os.system("start Help.html")
        except Exception as e:
            self.messageBox(str(e))

    def startAutocad(self):
        if self.path:
            def bukaAcad():
                try:
                    self.hide()
                    self.information.setText("Prosessing")
                    uiStart = Ui_progresWIndow(self.path)
                    uiStart.setModal(True)
                    uiStart.exec()
                    self.show()
                    self.information.setText("Drafting %s Selesai" % self.fileName)
                except Exception as e:
                    self.messageBox("Gagal Drafting", str(e) + "\n\nPeriksa Format Excel Interlocking Table")

            try:
                acad = Autocad()
                print(acad.doc.Name)
                bukaAcad()
            except:
                self.messageBox("AutoCad Belum Terbuka", "Klik OK\nLalu Tunggu Sampai AutoCad Terbuka Otomatis")
                try:
                    acad = Autocad(create_if_not_exists=True)
                    acad.prompt("Hello, Autocad from Python\n")
                    self.messageBox("Autocad Terbuka", "Lanjutkan Prosses")
                    bukaAcad()
                except Exception as e:
                    self.messageBox("Error", str(e) + "\n\nGagal Membuka AutoCad")
        else:
            self.messageBox("Gagal","Pilih File Excel Terlebih Dahulu")


    def messageBox(self, text, textLong):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText(text)
        msg.setInformativeText(textLong)
        msg.setWindowTitle("Information")
        msg.exec_()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    main = mainWindow()
    main.show()
    sys.exit(app.exec_())
