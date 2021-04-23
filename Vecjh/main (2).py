from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QWidget, QApplication, QPushButton
from PyQt5 import QtCore, uic, QtGui
from project.Sorting import create_itog

import sys

entry = []
count_cell = []
title_cell = []
btn = []

class Window_for_generate(QtWidgets.QMainWindow):

    def __init__(self):
        QWidget.__init__(self)
        self.set()

    def set(self):
        _translate = QtCore.QCoreApplication.translate
        self.w_root = uic.loadUi('Excel_table.ui')
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

        def create_entry(n):
            entry.append(QtWidgets.QLineEdit(self.w_root))
            entry[n].setGeometry(QtCore.QRect(55, 112 + 66 * n, 535, 28))
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
            entry[n].setMaxLength(150)
            entry[n].setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft)
            entry[n].setPlaceholderText('Адрес документа')

            btn.append(QPushButton(self.w_root))
            btn[n].setGeometry(610, 111+66*n, 50, 32)
            btn[n].setStyleSheet("QPushButton {height:48; background: rgba(0, 0, 0, 0.32); color: White; border-radius: 16px;padding-top: -15px;}"
                               "QPushButton:pressed {background-color:rgba(0, 0, 0, 0.52) ; }")
            btn[n].setText(_translate("MainWindow", "..."))
            btn[n].setFont(font_btn)
            btn[n].clicked.connect(lambda: insert_text())




            count_cell.append(QtWidgets.QLineEdit(self.w_root))
            count_cell[n].setGeometry(QtCore.QRect(680, 112 + 66 * n, 87, 28))
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
            title_cell[n].setGeometry(QtCore.QRect(782, 112 + 66 * n, 87, 28))
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

        num=5
        for n in range(num):
            create_entry(n)




        def read_files_path():
            files_df = []

            for n in range(num):
                if not entry[n].text() or not count_cell[n].text() or not title_cell[n].text():

                    continue
                files_df.append({'sourse': entry[n].text().replace('/', '\\'), 'num_yach': int(count_cell[n].text()), 'title_cell': title_cell[n].text()})
               #  files_df.append(entry[n].text())


            if files_df:
                return files_df
            else:
                files_df.append({'sourse': '', 'num_yach': '', 'title_cell': ''})
            # for n in range(5):
            #     print(entry[n].text())

        btn_itog = QPushButton(self.w_root)
        btn_itog.setGeometry(QtCore.QRect(55, 444, 300, 48))
        btn_itog.setStyleSheet(
            "QPushButton { background: rgba(0, 0, 0, 0.32); color: White; border-radius: 24px;}"
            "QPushButton:pressed {background-color:rgba(0, 0, 0, 0.52) ; }")
        btn_itog.setText(_translate("MainWindow", "Выполнить"))
        btn_itog.setFont(font_btn_itog)

        btn_itog.clicked.connect(lambda: create_itog(read_files_path()))




        self.w_root.show()





app = QApplication([sys.argv])
application = Window_for_generate()
sys.exit(app.exec_())
