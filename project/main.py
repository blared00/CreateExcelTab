from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QWidget, QApplication, QPushButton, QInputDialog, QMessageBox
from PyQt5 import QtCore, uic, QtGui
from project.Sorting import create_itog

import sys

from project.mydesigne import Ui_MainWindow

entry = []
count_cell = []
title_cell = []
btn = []


class Window_for_generate(QtWidgets.QMainWindow):

    def __init__(self):
        QWidget.__init__(self)

        self.showDialog()

    def showDialog(self):

        text, ok = QInputDialog.getText(self, 'Выбор количества типов ячеек',
                                        'Введите число типов ячеек:')

        if ok:
            try:
                self.sets(int(text))
            except:
                ret = QMessageBox.critical(self, "Ошибка ", "Необходимо ввести целое число\nколичества типов ячеек", QMessageBox.Ok)
                if ret==QMessageBox.Ok:
                    self.showDialog()

    def sets(self,num):
        _translate = QtCore.QCoreApplication.translate
        self.w_roots = uic.loadUi('resources/ui/Excel_table.ui')
        layout = QtWidgets.QVBoxLayout(self.w_roots)
        self.scrollArea = QtWidgets.QScrollArea(self.w_roots)
        self.scrollArea.setGeometry(0, 112, 954, 330)
        layout.addWidget(self.scrollArea)

        self.w_root = QtWidgets.QWidget()
        self.w_root.setGeometry(0, 0, 900, num*66)
        self.scrollArea.setWidget(self.w_root)

        font_btn = QtGui.QFont()
        font_btn.setFamily("Times New Roman")
        font_btn.setPointSize(28)
        font_btn.setBold(True)
        font_btn.setWeight(75)

        font_entry = QtGui.QFont()
        font_entry.setFamily("Arial")
        font_entry.setPointSize(16)

        font_btn_itog = QtGui.QFont()
        font_btn_itog.setFamily("Kharkiv Tone")
        font_btn_itog.setPointSize(25)

        def save_excel(Window):
            try:
                sources = read_files_path()
                if sources:
                    mes = create_itog(Window, sources)
                    if mes:
                        QMessageBox.information(Window, "Выполнено", "Сводная таблица успешно создана", QMessageBox.Ok)
            except Exception as e:
                print(e)

        def read_files_path():
            files_df = []

            for n in range(num):
                if not entry[n].text() or not count_cell[n].text() or not title_cell[n].text():
                    continue
                files_df.append({'sourse': entry[n].text().replace('/', '\\'), 'num_yach': int(count_cell[n].text()),
                                 'title_cell': title_cell[n].text()})
            #  files_df.append(entry[n].text())

            if files_df and len(files_df)>1:
                return files_df
            if len(files_df) == 1:
                QMessageBox.critical(self, "Ошибка ", "Выбран один Excel-фаил\nДля работы программы необходимо больше одного фаила", QMessageBox.Ok)
            else:
                QMessageBox.critical(self, "Ошибка ", "Некоторые поля не заполнены", QMessageBox.Ok)

        def create_entry(n):
            entry.append(QtWidgets.QLineEdit(self.w_root))
            entry[n].setGeometry(QtCore.QRect(55, 66 * n, 535, 28))
            sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            sizePolicy.setHorizontalStretch(0)
            sizePolicy.setVerticalStretch(0)
            sizePolicy.setHeightForWidth(entry[n].sizePolicy().hasHeightForWidth())
            entry[n].setSizePolicy(sizePolicy)
            entry[n].setFont(font_entry)
            entry[n].setStyleSheet("\n"
                                   "border-bottom: 2px solid white;\n"
                                   "background-color: #976EED;\n"
                                   "color: rgba(255, 255, 255, 1);\n"
                                   "")
            entry[n].setMaxLength(350)
            entry[n].setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft)
            entry[n].setPlaceholderText(f'{n+1}. Адрес документа')

            btn.append(QPushButton(self.w_root))
            btn[n].setGeometry(610, 66 * n, 50, 32)
            btn[n].setStyleSheet(
                "QPushButton {height:48; background: rgba(0, 0, 0, 0.32); color: White; border-radius: 16px;padding-top: -15px;}"
                "QPushButton:pressed {background-color:rgba(0, 0, 0, 0.52) ; }")
            btn[n].setText(_translate("MainWindow", "..."))
            btn[n].setFont(font_btn)
            btn[n].clicked.connect(lambda: insert_text())

            count_cell.append(QtWidgets.QLineEdit(self.w_root))
            count_cell[n].setGeometry(QtCore.QRect(680, 66 * n, 87, 28))
            sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            sizePolicy.setHorizontalStretch(0)
            sizePolicy.setVerticalStretch(0)
            sizePolicy.setHeightForWidth(count_cell[n].sizePolicy().hasHeightForWidth())
            count_cell[n].setSizePolicy(sizePolicy)
            count_cell[n].setFont(font_entry)
            count_cell[n].setStyleSheet("\n"
                                        "border-bottom: 2px solid white;\n"
                                        "background-color: #976EED;\n"
                                        "color: rgba(255, 255, 255, 0.82);\n"
                                        "")
            count_cell[n].setMaxLength(2)
            count_cell[n].setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft)
            count_cell[n].setPlaceholderText('кол-во')

            title_cell.append(QtWidgets.QLineEdit(self.w_root))
            title_cell[n].setGeometry(QtCore.QRect(782, 66 * n, 87, 28))
            sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            sizePolicy.setHorizontalStretch(0)
            sizePolicy.setVerticalStretch(0)
            sizePolicy.setHeightForWidth(title_cell[n].sizePolicy().hasHeightForWidth())
            title_cell[n].setSizePolicy(sizePolicy)
            title_cell[n].setFont(font_entry)
            title_cell[n].setStyleSheet("\n"
                                        "border-bottom: 2px solid white;\n"
                                        "background-color: #976EED;\n"
                                        "color: rgba(255, 255, 255, 0.82);\n"
                                        "")
            title_cell[n].setMaxLength(15)
            title_cell[n].setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft)
            title_cell[n].setPlaceholderText('обоз.яч')

            def insert_text():
                entry[n].clear()
                file, _ = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                'Выбор Excel-файла',
                                                                '',
                                                                'Excel-файлы (*.xlsx);;Все файлы(*.*)')
                entry[n].insert(file)

        # num = 5
        for n in range(num):
            create_entry(n)


            # for n in range(5):
            #     print(entry[n].text())

        btn_itog = QPushButton(self.w_roots)
        btn_itog.setGeometry(QtCore.QRect(55, 444, 300, 48))
        btn_itog.setStyleSheet(
            "QPushButton { background: rgba(0, 0, 0, 0.32); color: White; border-radius: 24px;}"
            "QPushButton:pressed {background-color:rgba(0, 0, 0, 0.52) ; }")
        btn_itog.setText(_translate("MainWindow", "Выполнить"))
        btn_itog.setFont(font_btn_itog)

        btn_itog.clicked.connect(lambda: save_excel(self))



        #     print('Все заебок')
        # else:
        #     print('Залупа')
        # if num >5:
        #     self.w_root.resize(914, 844)
        # self.w_root.center()
        self.setCentralWidget(self.centralWidget())
        self.w_roots.show()


Stylesheet_1 = ("""
QScrollBar:vertical {              
    border: none;
    background: white;
    width: 1px;               
    margin: 0px 0px 0px 0px;
}
QScrollBar::handle:vertical {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
    stop: 0 rgb(32, 47, 130), stop: 0.5 rgb(32, 47, 130), stop:1 rgb(32, 47, 130));
    min-height: 0px;
}
QScrollBar::add-line:vertical {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
    stop: 0 rgb(32, 47, 130), stop: 0.5 rgb(32, 47, 130),  stop:1 rgb(32, 47, 130));
    height: 0px;
    subcontrol-position: bottom;
    subcontrol-origin: margin;
}
QScrollBar::sub-line:vertical {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
    stop: 0  rgb(32, 47, 130), stop: 0.5 rgb(32, 47, 130),  stop:1 rgb(32, 47, 130));
    height: 0 px;
    subcontrol-position: top;
    subcontrol-origin: margin;
}
""")


app = QApplication([sys.argv])
app.setWindowIcon(QtGui.QIcon('resources/img/icon.ico'))
app.setStyleSheet(Stylesheet_1)

application = Window_for_generate()
application.w_roots.show()
sys.exit(app.exec_())
