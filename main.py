import sys
import sqlite3
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QTableWidgetItem, QMessageBox


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.con = sqlite3.connect("orders.sqlite")
        self.cur = self.con.cursor()
        self.filt = ''
        self.sorting_dict = {'По названию': lambda x: x[1],
                             'По дате создания': lambda x: x[0],
                             'По организации': lambda x: x[5],
                             'По дате': lambda x: x[3],
                             'По приказывающему': lambda x: x[8]}
        self.table = []
        uic.loadUi('mainWidget.ui', self)

        self.addButton.clicked.connect(self.add)
        self.pushButton.clicked.connect(self.filter)
        self.editButton.clicked.connect(self.edit)

    def add(self):
        self.addWindow = AddWindow(self)
        self.addWindow.show()

    def filter(self):
        self.filterWindow = FilterWindow(self)
        self.filterWindow.show()

    def edit(self):
        self.editWindow = EditWindow(self)
        self.editWindow.show()

    def update_table(self):
        try:
            self.table = list(self.cur.execute(f"""SELECT * FROM orders {self.filt}""").fetchall())
            self.listWidget.clear()
            self.table.sort(key=self.sorting_dict[self.comboBox.currentText()])
            for i in self.table:
                self.listWidget.addItem(i[1])
        except Exception as e:
            print(str(e))


class FilterWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('filterWidget.ui', self)


class AddWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        uic.loadUi('addWidget.ui', self)
        self.saveButton.clicked.connect(self.save)
        self.cancelButton.clicked.connect(self.close)
        self.addButton.clicked.connect(self.addRow)
        self.deleteButton.clicked.connect(self.delRow)
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.deleteButton.setEnabled(len(self.guysTable.selectedItems()) != 0))

    def save(self):
        try:
            guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                             for i in range(self.guysTable.rowCount())])
        except AttributeError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Пустые ячейки")
            msg.setText("Пустые ячейки в приказывающих")
            msg.exec_()
            return
        self.main_window.cur.execute(f"INSERT INTO orders(name,date,number,organization,title,reason,text,guys) VALUES"
                                     f"('{self.name.text()}', "
                                     f"'{self.date.date().toString('dd-MM-yyyy')}', "
                                     f"'{self.number.text()}', "
                                     f"'{self.organization.text()}', "
                                     f"'{self.title.text()}', "
                                     f"'{self.reason.toPlainText()}', "
                                     f"'{self.text.toPlainText()}', "
                                     f"'{guys}')")
        self.main_window.update_table()
        self.close()

    def addRow(self):
        self.guysTable.setRowCount(self.guysTable.rowCount() + 1)
        for i in range(2):
            self.guysTable.setItem(self.guysTable.rowCount(), i, QTableWidgetItem())

    def delRow(self):
        selected_items = self.guysTable.selectedItems()
        for item in selected_items:
            self.guysTable.removeRow(item.row())


class EditWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('editWidget.ui', self)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
