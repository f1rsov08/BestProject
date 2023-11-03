import sys
import sqlite3
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QTableWidgetItem, QMessageBox
from PyQt5.QtCore import QDate


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.con = sqlite3.connect("orders.sqlite")
        self.cur = self.con.cursor()
        self.filt = ''
        self.sorting_dict = {'По названию': lambda x: self.table[x][0],
                             'По дате': lambda x: self.table[x][1],
                             'По номеру': lambda x: self.table[x][2],
                             'По организации': lambda x: self.table[x][3],
                             'По заголовку': lambda x: self.table[x][4],
                             'По основанию': lambda x: self.table[x][5],
                             'По тексту': lambda x: self.table[x][6]}
        self.table = {}
        uic.loadUi('mainWidget.ui', self)

        self.addButton.clicked.connect(self.add)
        self.filterButton.clicked.connect(self.filter)
        self.searchButton.clicked.connect(self.update_table)
        self.listWidget.itemDoubleClicked.connect(self.edit)
        self.update_table()

    def add(self):
        self.addWindow = AddWindow(self)
        self.addWindow.show()

    def filter(self):
        self.filterWindow = FilterWindow(self)
        self.filterWindow.show()

    def edit(self, item):
        self.editWindow = EditWindow(self, item)
        self.editWindow.show()

    def update_table(self):
        self.table = {i[0]: i[1:] for i in self.cur.execute(f"""SELECT * FROM orders {self.filt}""")}
        self.listWidget.clear()
        for i in sorted(self.table.keys(), key=self.sorting_dict[self.comboBox.currentText()]):
            self.listWidget.addItem(i)


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
        self.addRowButton.clicked.connect(self.addRow)
        self.delRowButton.clicked.connect(self.delRow)
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.delRowButton.setEnabled(len(self.guysTable.selectedItems()) != 0))
        self.name.textChanged.connect(self.check_input)

    def check_input(self):
        if self.name.text() in self.main_window.table.keys():
            self.saveButton.setDisabled(True)
            self.statusbar.setText('Название уже занято')
        elif not self.name.text():
            self.saveButton.setDisabled(True)
            self.statusbar.setText('Название не может быть пустым')
        else:
            self.saveButton.setEnabled(True)
            self.statusbar.setText('')

    def save(self):
        try:
            guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                             for i in range(self.guysTable.rowCount())])
        except AttributeError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Пустые ячейки")
            msg.setText("Пустые ячейки в приказывающих")
        else:
            reason, text = self.reason.toPlainText().replace('\n', '\\n'), self.text.toPlainText().replace('\n', '\\n')
            self.main_window.cur.execute(
                f"INSERT INTO orders(name,date,number,organization,title,reason,text,guys) VALUES"
                f"('{self.name.text()}', "
                f"'{self.date.date().toString('dd-MM-yyyy')}', "
                f"'{self.number.text()}', "
                f"'{self.organization.text()}', "
                f"'{self.title.text()}', "
                f"'{reason}', "
                f"'{text}', "
                f"'{guys}')")
            self.main_window.con.commit()
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
    def __init__(self, main_window, item):
        super().__init__()
        self.item = item
        self.main_window = main_window
        uic.loadUi('editWidget.ui', self)
        self.saveButton.clicked.connect(self.save)
        self.delRowButton.clicked.connect(self.delRow)
        self.cancelButton.clicked.connect(self.close)
        self.addRowButton.clicked.connect(self.addRow)
        self.deleteButton.clicked.connect(self.delete)
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.delRowButton.setEnabled(len(self.guysTable.selectedItems()) != 0))
        self.name.textChanged.connect(self.check_input)
        self.guysTable.itemChanged.connect(self.check_input)

        self.name.setText(self.item.text())
        day, month, year = map(int, self.main_window.table[self.item.text()][0].split('-'))
        self.date.setDate(QDate(year, month, day))
        self.number.setValue(int(self.main_window.table[self.item.text()][1]))
        self.title.setText(self.main_window.table[self.item.text()][2])
        self.organization.setText(self.main_window.table[self.item.text()][3])
        self.reason.setPlainText(self.main_window.table[self.item.text()][4].replace('\\n', '\n'))
        self.text.setPlainText(self.main_window.table[self.item.text()][5].replace('\\n', '\n'))

        if self.main_window.table[self.item.text()][6]:
            self.guysTable.setRowCount(len(self.main_window.table[self.item.text()][6].split(',')))
            for i, row in enumerate(self.main_window.table[self.item.text()][6].split(',')):
                for j, elem in enumerate(row.split(':')):
                    self.guysTable.setItem(i, j, QTableWidgetItem(elem))

    def check_input(self):
        if not self.name.text():
            self.saveButton.setDisabled(True)
            self.statusbar.setText('Название не может быть пустым')
        elif self.name.text() in self.main_window.table.keys() and self.name.text() != self.item.text():
            self.saveButton.setDisabled(True)
            self.statusbar.setText('Название уже занято')
        else:
            self.saveButton.setEnabled(True)
            self.statusbar.setText('')

    def save(self):
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        reason, text = self.reason.toPlainText().replace('\n', '\\n'), self.text.toPlainText().replace('\n', '\\n')
        self.main_window.cur.execute(
            f"UPDATE orders "
            f"SET name = '{self.name.text()}', "
            f"date = '{self.date.date().toString('dd-MM-yyyy')}', "
            f"number = '{self.number.text()}', "
            f"organization = '{self.organization.text()}', "
            f"title = '{self.title.text()}', "
            f"reason = '{reason}', "
            f"text = '{text}', "
            f"guys = '{guys}' "
            f"WHERE name = '{self.item.text()}'")
        self.main_window.con.commit()
        self.main_window.update_table()
        self.close()

    def delete(self):
        valid = QMessageBox.question(
            self, '', f"Вы действительно хотите безвозвратно удалить {self.item.text()}?",
            QMessageBox.Yes, QMessageBox.No)
        if valid == QMessageBox.Yes:
            self.main_window.cur.execute(f"DELETE from orders where name = '{self.item.text()}'")
            self.main_window.con.commit()
            self.main_window.update_table()
            self.close()

    def addRow(self):
        self.guysTable.setRowCount(self.guysTable.rowCount() + 1)
        for i in range(2):
            self.guysTable.setItem(self.guysTable.rowCount(), i, QTableWidgetItem())
        self.replace_empty_cells()

    def delRow(self):
        selected_items = self.guysTable.selectedItems()
        for item in selected_items:
            self.guysTable.removeRow(item.row())

    def replace_empty_cells(self):
        for i in range(self.guysTable.columnCount()):
            for j in range(self.guysTable.rowCount()):
                if self.guysTable.item(j, i) is None:
                    self.guysTable.setItem(j, i, QTableWidgetItem(''))
        return False


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
