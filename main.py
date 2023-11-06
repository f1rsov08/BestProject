import sys
import sqlite3

from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QTableWidgetItem, QMessageBox, QFileDialog
from PyQt5.QtCore import QDate

from docxtpl import DocxTemplate


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.con = sqlite3.connect("orders.sqlite")
        self.cur = self.con.cursor()
        self.sorting_dict = {'По названию': lambda x: x,
                             'По дате': lambda x: self.table[x][0],
                             'По номеру': lambda x: self.table[x][1],
                             'По организации': lambda x: self.table[x][2],
                             'По заголовку': lambda x: self.table[x][3],
                             'По основанию': lambda x: self.table[x][4],
                             'По тексту': lambda x: self.table[x][5]}
        self.table = {}
        uic.loadUi('ui-files/mainWidget.ui', self)

        self.comboBox.activated.connect(self.update_table)
        self.addButton.clicked.connect(self.add)
        self.searchButton.clicked.connect(self.update_table)
        self.listWidget.itemDoubleClicked.connect(self.edit)
        self.update_table()

    def add(self):
        self.addWindow = AddWindow(self)
        self.addWindow.show()

    def edit(self, item):
        self.editWindow = EditWindow(self, item)
        self.editWindow.show()

    def update_table(self):
        if self.search.text():
            self.table = {i[0]: i[1:] for i in self.cur.execute(f"""SELECT * FROM orders 
            WHERE name LIKE '%{self.search.text()}%'""")}
        else:
            self.table = {i[0]: i[1:] for i in self.cur.execute(f"""SELECT * FROM orders""")}
        self.listWidget.clear()
        for i in sorted(self.table.keys(), key=self.sorting_dict[self.comboBox.currentText()]):
            self.listWidget.addItem(i)


class AddWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        uic.loadUi('ui-files/addWidget.ui', self)
        self.saveButton.clicked.connect(self.save)
        self.cancelButton.clicked.connect(self.close)
        self.addRowButton.clicked.connect(self.addRow)
        self.delRowButton.clicked.connect(self.delRow)
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.delRowButton.setEnabled(len(self.guysTable.selectedItems()) != 0))
        self.name.textChanged.connect(self.check_input)

        self.name.textChanged.connect(self.progress_bar_edit)
        self.number.valueChanged.connect(self.progress_bar_edit)
        self.title.textChanged.connect(self.progress_bar_edit)
        self.organization.textChanged.connect(self.progress_bar_edit)
        self.reason.textChanged.connect(self.progress_bar_edit)
        self.text.textChanged.connect(self.progress_bar_edit)

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
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        reason, text = self.reason.toPlainText().replace('\n', '\\n'), self.text.toPlainText().replace('\n', '\\n')

        self.main_window.cur.execute(f"INSERT INTO organizations (name) VALUES ('{self.organization.text()}')")
        organization_id = self.main_window.cur.lastrowid

        self.main_window.cur.execute(
            f"INSERT INTO orders(name,date,number,organization,title,reason,text,guys) VALUES"
            f"('{self.name.text()}', "
            f"'{self.date.date().toString('yyyy-MM-dd')}', "
            f"'{self.number.text()}', "
            f"{organization_id}, "
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
        self.replace_empty_cells()
        self.progress_bar_edit()

    def delRow(self):
        selected_items = self.guysTable.selectedItems()
        for item in selected_items:
            self.guysTable.removeRow(item.row())
        self.progress_bar_edit()

    def replace_empty_cells(self):
        for i in range(self.guysTable.columnCount()):
            for j in range(self.guysTable.rowCount()):
                if self.guysTable.item(j, i) is None:
                    self.guysTable.setItem(j, i, QTableWidgetItem(''))
        return False

    def progress_bar_edit(self):
        self.progressBar.setValue((bool(self.name.text()) +
                                   bool(self.number.value()) +
                                   bool(self.organization.text()) +
                                   bool(self.title.text()) +
                                   bool(self.reason.toPlainText()) +
                                   bool(self.text.toPlainText()) +
                                   bool(self.guysTable.rowCount()) +
                                   1) * 125)


class EditWindow(QWidget):
    def __init__(self, main_window, item):
        super().__init__()
        self.item = item
        self.main_window = main_window

        self.organization_id = self.main_window.table[self.item.text()][2]

        uic.loadUi('ui-files/editWidget.ui', self)
        self.saveButton.clicked.connect(self.save)
        self.delRowButton.clicked.connect(self.delRow)
        self.cancelButton.clicked.connect(self.close)
        self.addButton.clicked.connect(self.add)
        self.addRowButton.clicked.connect(self.addRow)
        self.deleteButton.clicked.connect(self.delete)
        self.exportButton.clicked.connect(self.export)
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.delRowButton.setEnabled(len(self.guysTable.selectedItems()) != 0))
        self.name.textChanged.connect(self.check_input)
        self.guysTable.itemChanged.connect(self.check_input)

        org_name = self.main_window.cur.execute(f"SELECT name "
                                                f"FROM organizations "
                                                f"WHERE id = '{self.organization_id}'").fetchone()[0]

        self.name.setText(self.item.text())
        year, month, day = map(int, self.main_window.table[self.item.text()][0].split('-'))
        self.date.setDate(QDate(year, month, day))
        self.number.setValue(int(self.main_window.table[self.item.text()][1]))
        self.title.setText(self.main_window.table[self.item.text()][3])
        self.organization.setText(org_name)
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
            self.addButton.setDisabled(True)
            self.statusbar.setText('Название не может быть пустым')
        elif self.name.text() in self.main_window.table.keys() and self.name.text() != self.item.text():
            self.saveButton.setDisabled(True)
            self.addButton.setDisabled(True)
            self.statusbar.setText('Название уже занято')
        elif self.name.text() in self.main_window.table.keys() and self.name.text() == self.item.text():
            self.saveButton.setEnabled(True)
            self.addButton.setDisabled(True)
        else:
            self.saveButton.setEnabled(True)
            self.addButton.setEnabled(True)
            self.statusbar.setText('')

    def save(self):
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        reason, text = self.reason.toPlainText().replace('\n', '\\n'), self.text.toPlainText().replace('\n', '\\n')

        self.main_window.cur.execute(f"UPDATE organizations "
                                     f"SET name = '{self.organization.text()}' "
                                     f"WHERE id = {self.organization_id}")

        self.main_window.cur.execute(
            f"UPDATE orders "
            f"SET name = '{self.name.text()}', "
            f"date = '{self.date.date().toString('yyyy-MM-dd')}', "
            f"number = '{self.number.text()}', "
            f"organization = {self.organization_id}, "
            f"title = '{self.title.text()}', "
            f"reason = '{reason}', "
            f"text = '{text}', "
            f"guys = '{guys}' "
            f"WHERE name = '{self.item.text()}'")

        self.main_window.con.commit()
        self.main_window.update_table()
        self.close()

    def add(self):
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        reason, text = self.reason.toPlainText().replace('\n', '\\n'), self.text.toPlainText().replace('\n', '\\n')

        self.main_window.cur.execute(f"INSERT INTO organizations (name) VALUES ('{self.organization.text()}')")
        self.main_window.con.commit()
        organization_id = self.main_window.cur.lastrowid

        self.main_window.cur.execute(
            f"INSERT INTO orders(name,date,number,organization,title,reason,text,guys) VALUES"
            f"('{self.name.text()}', "
            f"'{self.date.date().toString('yyyy-MM-dd')}', "
            f"'{self.number.text()}', "
            f"{organization_id}, "
            f"'{self.title.text()}', "
            f"'{reason}', "
            f"'{text}', "
            f"'{guys}')")
        self.main_window.con.commit()
        self.main_window.update_table()
        self.close()

    def delete(self):
        valid = QMessageBox.question(
            self, '', f"Вы действительно хотите безвозвратно удалить {self.item.text()}?",
            QMessageBox.Yes, QMessageBox.No)
        if valid == QMessageBox.Yes:
            self.main_window.cur.execute(f"DELETE from orders where name = '{self.item.text()}'")
            self.main_window.cur.execute(f"DELETE from organizations where id = {self.organization_id}")

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

    def export(self):
        try:
            fname = QFileDialog.getSaveFileName(
                self, 'Сохранение файла', '',
                'Документ Word (*.docx);;Текстовые документы (*.txt);;Все файлы (*)')[0]
            month_dict = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября',
                          'октября', 'ноября', 'декабря']
            if fname[-4:] == '.txt':
                guys = '\n'.join([f'{self.guysTable.item(i, 0).text()}\t{self.guysTable.item(i, 1).text()}'
                                  for i in range(self.guysTable.rowCount())])
                with open(fname, 'w', encoding='utf8') as file:
                    file.write(f'''{self.organization.text()}

ПРИКАЗ

{self.date.date().day()} {month_dict[self.date.date().month()]} {self.date.date().year()} года.\t№{self.number.text()}


{self.title.text()}


{self.reason.toPlainText()}

ПРИКАЗЫВАЮ:
{self.text.toPlainText()}

{guys}
''')
            elif fname[-5:] == '.docx':
                doc = DocxTemplate('template.docx')
                space_len = lambda x: 47 - len(self.guysTable.item(x, 0).text() + self.guysTable.item(x, 1).text())
                context = {'organization': self.organization.text(),
                           'date': f'{self.date.date().day()} '
                                   f'{month_dict[self.date.date().month()]} '
                                   f'{self.date.date().year()} года',
                           'num': self.number.text(),
                           'title': self.title.text(),
                           'reason': self.reason.toPlainText().split('\n'),
                           'text': self.text.toPlainText().split('\n'),
                           'items': [f'{self.guysTable.item(i, 0).text()}'
                                     f'{" " * space_len(i)}'
                                     f'{self.guysTable.item(i, 1).text()}'
                                     for i in range(self.guysTable.rowCount())]}
                doc.render(context)
                doc.save(fname)
                finish = FinishWidget()
                finish.show()
        except PermissionError:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("Внимание")
            msg_box.setText("Файл открыт в другой программе.")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()


class FinishWidget(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui-files/finishWidget.ui', self)
        self.pushButton.clicked.connect(self.close)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
