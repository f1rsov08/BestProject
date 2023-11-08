import sys
import sqlite3

from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QTableWidgetItem, QMessageBox, QFileDialog
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QPixmap, QIcon

import webbrowser

from docxtpl import DocxTemplate


class MainWindow(QMainWindow):
    """
    Главное окно
    """

    def __init__(self):
        """
        Инициализация главного окна
        """
        super().__init__()

        # Подключение к базе данных
        self.con = sqlite3.connect("orders.sqlite")
        self.cur = self.con.cursor()

        # Создание словаря с ключами для сортировки
        self.sorting_dict = {'По названию': lambda x: x,
                             'По дате': lambda x: self.table[x][0],
                             'По номеру': lambda x: self.table[x][1],
                             'По организации': lambda x: self.table[x][2],
                             'По заголовку': lambda x: self.table[x][3],
                             'По основанию': lambda x: self.table[x][4],
                             'По тексту': lambda x: self.table[x][5]}

        # Создание словаря для хранения базы данных
        self.table = {}

        # Дизайн
        uic.loadUi('ui-files/mainWidget.ui', self)
        self.setStyleSheet('.QWidget {background-image: url("images/background.jpg");}')
        self.centralwidget.setContentsMargins(10, 10, 10, 10)

        # Подключение
        # Подключение кнопки открытия окна добавления приказа к функции открытия окна добавления приказов
        self.addButton.clicked.connect(self.add)
        # Подключение кнопки поиска к функции обновления списка приказов
        self.searchButton.clicked.connect(self.update_list)
        # Подключение кнопки открытия окна с информацией о программе к функции открытия окна с информацией о программе
        self.aboutButton.clicked.connect(self.about)
        # Подключение выпадающего списка к функции обновления списка приказов
        self.comboBox.activated.connect(self.update_list)
        # Подключение двойного клика по элементу виджета списка приказов к функции открытия окна редактирования приказов
        self.listWidget.itemDoubleClicked.connect(self.edit)

        # Обновление списка приказов
        self.update_list()

    def add(self):
        """
        Функция открытия окна добавления приказов
        """
        # Создание окна добавления приказов
        self.addWindow = AddWindow(self)
        # Открытие окна добавления приказов
        self.addWindow.show()

    def edit(self, item):
        """
        Функция открытия окна редактирования приказов
        """
        # Создание окна редактирования приказов
        self.editWindow = EditWindow(self, item)
        # Открытие окна редактирования приказов
        self.editWindow.show()

    def about(self):
        """
        Функция открытия окна с информацией о программе
        """
        # Создание окна с информацией о программе
        self.aboutWindow = AboutWindow()
        # Открытие окна с информацией о программе
        self.aboutWindow.show()

    def update_list(self):
        """
        Функция обновления списка приказов
        """
        # Если есть текст поиска
        if self.search.text():
            # Отправляем запрос к базе данных для получения приказов, где есть текст поиска
            # создаем словарь с ключами - названиями приказов
            # и значениями - списками с остальными элементами приказа (дата, номер и тд.)
            self.table = {i[0]: i[1:] for i in self.cur.execute(f"""SELECT * FROM orders 
            WHERE name LIKE '%{self.search.text()}%'""")}
        else:
            # Отправляем запрос к базе данных для получения всех приказов
            # создаем словарь с ключами - названиями приказов
            # и значениями - списками с остальными элементами приказа (дата, номер и тд.)
            self.table = {i[0]: i[1:] for i in self.cur.execute(f"""SELECT * FROM orders""")}
        # Очищаем виджет списка
        self.listWidget.clear()
        # Добавляем в виджет списка отсортированные названия приказов
        for i in sorted(self.table.keys(), key=self.sorting_dict[self.comboBox.currentText()]):
            self.listWidget.addItem(i)


class AddWindow(QWidget):
    """
    Окно добавления приказов
    """

    def __init__(self, main_window):
        """
        Инициализация окна добавления приказов
        """
        super().__init__()
        self.main_window = main_window

        # Дизайн
        uic.loadUi('ui-files/addWidget.ui', self)
        self.setStyleSheet('.QWidget {background-image: url("images/background.jpg");}')

        # Подключение кнопок
        # Подключение кнопки сохранения приказа к функции сохранения приказа
        self.saveButton.clicked.connect(self.save)
        # Подключение кнопки отмены к функции закрытия окна
        self.cancelButton.clicked.connect(self.close)
        # Подключение кнопки добавления приказывающего к функции добавления приказывающего
        self.addRowButton.clicked.connect(self.addRow)
        # Подключение кнопки удаления приказывающего к функции удаления приказывающего
        self.delRowButton.clicked.connect(self.delRow)

        # Разблокировка кнопки удаления приказывающего, только если он выбран
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.delRowButton.setEnabled(len(self.guysTable.selectedItems()) != 0))

        # Подключение формы для ввода названия приказа к функции проверки названия
        self.name.textChanged.connect(self.check_name)

        # Подключение всех форм к функции для изменения прогресс бара
        # Подключение формы для ввода названия приказа
        self.name.textChanged.connect(self.progress_bar_edit)
        # Подключение формы для ввода номера приказа
        self.number.valueChanged.connect(self.progress_bar_edit)
        # Подключение формы для ввода заголовка приказа
        self.title.textChanged.connect(self.progress_bar_edit)
        # Подключение формы для ввода организации приказа
        self.organization.textChanged.connect(self.progress_bar_edit)
        # Подключение формы для ввода основания приказа
        self.reason.textChanged.connect(self.progress_bar_edit)
        # Подключение формы для ввода текста приказа
        self.text.textChanged.connect(self.progress_bar_edit)
        # Мы не подключили форму для ввода даты приказа, т.к. она всегда введена

    def check_name(self):
        """
        Функция проверки названия
        """
        # Если название пустое
        if not self.name.text():
            # Блокируем кнопку сохранения
            self.saveButton.setDisabled(True)
            # Выводим в статусбар сообщение, что название не может быть пустым
            self.statusbar.setText('Название не может быть пустым')
        # Если название уже используется
        elif self.name.text() in self.main_window.table.keys():
            # Блокируем кнопку сохранения
            self.saveButton.setDisabled(True)
            # Выводим в страусбар сообщение, что название уже занято
            self.statusbar.setText('Название уже занято')
        # Если все ок
        else:
            # Разблокируем кнопку сохранения
            self.saveButton.setEnabled(True)
            # Очищаем статусбар
            self.statusbar.setText('')

    def save(self):
        """
        Функция сохранения приказа
        """
        # Приводим приказывающих к формату "должность1:фио1,должность2:фио2,..."
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        # Меняем перенос строки на \n для хранения в базе данных в основании приказа
        reason = self.reason.toPlainText().replace('\n', '\\n')
        # Меняем перенос строки на \n для хранения в базе данных в тексте приказа
        text = self.text.toPlainText().replace('\n', '\\n')
        # Добавляем организацию в базу данных
        self.main_window.cur.execute(f"INSERT INTO organizations (name) VALUES ('{self.organization.text()}')")

        # Извлечение данных приказа
        # Извлекаем айди организации
        organization_id = self.main_window.cur.lastrowid
        # Извлекаем название приказа
        name = self.name.text()
        # Извлекаем дату приказа
        date = self.date.date().toString('yyyy-MM-dd')
        # Извлекаем номер приказа
        number = self.number.text()
        # Извлекаем заголовок приказа
        title = self.title.text()

        # Добавляем приказ в базу данных
        self.main_window.cur.execute(
            f"INSERT INTO orders(name,date,number,organization,title,reason,text,guys) VALUES"
            f"('{name}', "
            f"'{date}', "
            f"'{number}', "
            f"{organization_id}, "
            f"'{title}', "
            f"'{reason}', "
            f"'{text}', "
            f"'{guys}')")

        # Сохраняем изменения
        self.main_window.con.commit()
        # Обновляем список в главном окне
        self.main_window.update_list()
        # Закрываем окно
        self.close()

    def addRow(self):
        """
        Функция добавления приказывающего
        """
        # Добавление ряда для нового приказывающего
        self.guysTable.setRowCount(self.guysTable.rowCount() + 1)
        # Запись номера ряда приказывающего в таблице
        rowNum = self.guysTable.rowCount() - 1
        # Замена ячеек с None на пустые str
        for i in range(2):
            self.guysTable.setItem(rowNum, i, QTableWidgetItem(''))
        # Изменение прогресс бара
        self.progress_bar_edit()

    def delRow(self):
        """
        Функция удаления приказывающего
        """
        # Запись выбранных ячеек таблицы
        selected_items = self.guysTable.selectedItems()
        # Удаление выбранных ячеек таблицы
        for i in selected_items:
            self.guysTable.removeRow(i.row())
        # Изменение прогресс бара
        self.progress_bar_edit()

    def progress_bar_edit(self):
        """
        Функиця изменения прогресс бара
        """
        # Выставляем в прогресс бар количество заполненных форм умноженное на 125
        # (при всех заполненных формах 8 * 125 = 1000 = максимальное значение прогресс бара)
        # (дата всегда заполнена, поэтому она здесь заменена на +1)
        self.progressBar.setValue((bool(self.name.text()) +
                                   bool(self.number.value()) +
                                   bool(self.organization.text()) +
                                   bool(self.title.text()) +
                                   bool(self.reason.toPlainText()) +
                                   bool(self.text.toPlainText()) +
                                   bool(self.guysTable.rowCount()) +
                                   1) * 125)


class EditWindow(QWidget):
    """
    Окно редактирования приказа
    """

    def __init__(self, main_window, item):
        """
        Инициализация окна редактирования приказа
        """
        super().__init__()
        self.item = item
        self.main_window = main_window

        # Запись id организации
        self.organization_id = self.main_window.table[self.item.text()][2]

        # Дизайн
        uic.loadUi('ui-files/editWidget.ui', self)
        self.setStyleSheet('.QWidget {background-image: url("images/background.jpg");}')

        # Подключение кнопок
        # Подключение кнопки сохранения приказа к функции сохранения приказа
        self.saveButton.clicked.connect(self.save)
        # Подключение кнопки отмены к функции закрытия окна
        self.cancelButton.clicked.connect(self.close)
        # Подключение кнопки создания нового приказа к функции создания нового приказа
        self.addButton.clicked.connect(self.add)
        # Подключение кнопки добавления приказывающего к функции добавления приказывающего
        self.addRowButton.clicked.connect(self.addRow)
        # Подключение кнопки удаления приказывающего к функции удаления приказывающего
        self.delRowButton.clicked.connect(self.delRow)
        # Подключение кнопки удаления приказа к функции удаления приказа
        self.deleteButton.clicked.connect(self.delete)
        # Подключение кнопки экспорта к функции экспорта
        self.exportButton.clicked.connect(self.export)

        # Разблокировка кнопки удаления приказывающего, только если он выбран
        self.guysTable.itemSelectionChanged.connect(
            lambda: self.delRowButton.setEnabled(len(self.guysTable.selectedItems()) != 0))

        # Подключение формы для ввода названия приказа к функции проверки названия
        self.name.textChanged.connect(self.check_name)

        # Извлечение данных приказа
        # Извлечение названия приказа
        name = self.item.text()
        # Извлечение даты приказа
        year, month, day = map(int, self.main_window.table[self.item.text()][0].split('-'))
        # Извлечение номера приказа
        number = int(self.main_window.table[self.item.text()][1])
        # Извлечение заголовка приказа
        title = self.main_window.table[self.item.text()][3]
        # Извлечение названия организации
        organization_name = self.main_window.cur.execute(f"SELECT name "
                                                         f"FROM organizations "
                                                         f"WHERE id = '{self.organization_id}'").fetchone()[0]
        # Извлечение основания приказа
        reason = self.main_window.table[self.item.text()][4].replace('\\n', '\n')
        # Извлечение текста приказа
        text = self.main_window.table[self.item.text()][5].replace('\\n', '\n')
        # Извлечение приказывающих
        guys = self.main_window.table[self.item.text()][6].split(',')

        # Заполнение форм данными приказа
        # Заполнение формы для ввода названия приказа
        self.name.setText(name)
        # Заполнение формы для ввода даты приказа
        self.date.setDate(QDate(year, month, day))
        # Заполнение формы для ввода номера приказа
        self.number.setValue(number)
        # Заполнение формы для ввода заголовка приказа
        self.title.setText(title)
        # Заполнение формы для ввода организации
        self.organization.setText(organization_name)
        # Заполнение формы для ввода основания приказа
        self.reason.setPlainText(reason)
        # Заполнение формы для ввода текста приказа
        self.text.setPlainText(text)
        # Заполнение формы для ввода приказывающих, если они есть
        if guys != ['']:
            # Устанавливаем кол-во приказывающих
            self.guysTable.setRowCount(len(guys))
            # Устанавливаем приказывающих
            for i, row in enumerate(guys):
                for j, elem in enumerate(row.split(':')):
                    self.guysTable.setItem(i, j, QTableWidgetItem(elem))

    def check_name(self):
        """
        Функция проверки названия
        """
        # Если название пустое
        if not self.name.text():
            # Блокируем кнопку сохранения
            self.saveButton.setDisabled(True)
            # Блокируем кнопку создания нового приказа
            self.addButton.setDisabled(True)
            # Выводим в страусбар сообщение, что название не может быть пустым
            self.statusbar.setText('Название не может быть пустым')
        # Если название уже используется, но не редактируемым приказом
        elif self.name.text() in self.main_window.table.keys() and self.name.text() != self.item.text():
            # Блокируем кнопку сохранения
            self.saveButton.setDisabled(True)
            # Блокируем кнопку создания нового приказа
            self.addButton.setDisabled(True)
            # Выводим в страусбар сообщение, что название уже занято
            self.statusbar.setText('Название уже занято')
        # Если название уже используется редактируемым приказом
        elif self.name.text() in self.main_window.table.keys() and self.name.text() == self.item.text():
            # Разблокируем кнопку сохранения
            self.saveButton.setEnabled(True)
            # Блокируем кнопку создания нового приказа
            self.addButton.setDisabled(True)
        # Если все ок
        else:
            # Разблокируем кнопку сохранения
            self.saveButton.setEnabled(True)
            # Разблокируем кнопку создания нового приказа
            self.addButton.setEnabled(True)
            # Очищаем статусбар
            self.statusbar.setText('')

    def save(self):
        """
        Функция сохранения приказа
        """
        # Приводим приказывающих к формату "должность1:фио1,должность2:фио2,..."
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        # Меняем перенос строки на \n для хранения в базе данных в основании приказа
        reason = self.reason.toPlainText().replace('\n', '\\n')
        # Меняем перенос строки на \n для хранения в базе данных в тексте приказа
        text = self.text.toPlainText().replace('\n', '\\n')

        # Обновляем организацию в базе данных
        self.main_window.cur.execute(f"UPDATE organizations "
                                     f"SET name = '{self.organization.text()}' "
                                     f"WHERE id = {self.organization_id}")

        # Извлечение данных приказа
        # Извлекаем название приказа
        name = self.name.text()
        # Извлекаем дату приказа
        date = self.date.date().toString('yyyy-MM-dd')
        # Извлекаем номер приказа
        number = self.number.text()
        # Извлекаем заголовок приказа
        title = self.title.text()
        # Извлекаем старое название приказа
        item = self.item.text()

        # Изменяем приказ в базе данных
        self.main_window.cur.execute(
            f"UPDATE orders "
            f"SET name = '{name}', "
            f"date = '{date}', "
            f"number = '{number}', "
            f"organization = {self.organization_id}, "
            f"title = '{title}', "
            f"reason = '{reason}', "
            f"text = '{text}', "
            f"guys = '{guys}' "
            f"WHERE name = '{item}'")

        # Сохраняем изменения
        self.main_window.con.commit()
        # Обновляем список в главном окне
        self.main_window.update_list()
        # Закрываем окно
        self.close()

    def add(self):
        """
        Функция создания нового приказа из старого
        """
        # Приводим приказывающих к формату "должность1:фио1,должность2:фио2,..."
        guys = ','.join([f'{self.guysTable.item(i, 0).text()}:{self.guysTable.item(i, 1).text()}'
                         for i in range(self.guysTable.rowCount())])
        # Меняем перенос строки на \n для хранения в базе данных в основании приказа
        reason = self.reason.toPlainText().replace('\n', '\\n')
        # Меняем перенос строки на \n для хранения в базе данных в тексте приказа
        text = self.text.toPlainText().replace('\n', '\\n')

        # Обновляем организацию в базе данных
        self.main_window.cur.execute(f"INSERT INTO organizations (name) VALUES ('{self.organization.text()}')")

        # Извлечение данных приказа
        # Извлекаем номер организации
        organization_id = self.main_window.cur.lastrowid
        # Извлекаем название приказа
        name = self.name.text()
        # Извлекаем дату приказа
        date = self.date.date().toString('yyyy-MM-dd')
        # Извлекаем номер приказа
        number = self.number.text()
        # Извлекаем заголовок приказа
        title = self.title.text()

        # Добавляем приказ в базу данных
        self.main_window.cur.execute(
            f"INSERT INTO orders(name,date,number,organization,title,reason,text,guys) VALUES"
            f"('{name}', "
            f"'{date}', "
            f"'{number}', "
            f"{organization_id}, "
            f"'{title}', "
            f"'{reason}', "
            f"'{text}', "
            f"'{guys}')")

        # Сохраняем изменения
        self.main_window.con.commit()
        # Обновляем список в главном окне
        self.main_window.update_list()
        # Закрываем окно
        self.close()

    def delete(self):
        """
        Функция удаления приказа
        """
        # Спрашиваем у пользователя, точно ли он хочет удалить приказ
        valid = QMessageBox.question(
            self, '', f"Вы действительно хотите безвозвратно удалить {self.item.text()}?",
            QMessageBox.Yes, QMessageBox.No)
        # Если он соглашается
        if valid == QMessageBox.Yes:
            # Удаляем организацию
            self.main_window.cur.execute(f"DELETE from orders where name = '{self.item.text()}'")
            # Удаляем приказ
            self.main_window.cur.execute(f"DELETE from organizations where id = {self.organization_id}")

            # Сохраняем изменения
            self.main_window.con.commit()
            # Обновляем список в главном окне
            self.main_window.update_list()
            # Закрываем окно
            self.close()

    def addRow(self):
        """
        Функция добавления приказывающего
        """
        # Добавление ряда для нового приказывающего
        self.guysTable.setRowCount(self.guysTable.rowCount() + 1)
        # Запись номера ряда приказывающего в таблице
        rowNum = self.guysTable.rowCount() - 1
        # Замена ячеек с None на пустые str
        for i in range(2):
            self.guysTable.setItem(rowNum, i, QTableWidgetItem(''))
        # Изменение прогресс бара
        self.progress_bar_edit()

    def delRow(self):
        """
        Функция удаления приказывающего
        """
        # Запись выбранных ячеек таблицы
        selected_items = self.guysTable.selectedItems()
        # Удаление выбранных ячеек таблицы
        for i in selected_items:
            self.guysTable.removeRow(i.row())
        # Изменение прогресс бара
        self.progress_bar_edit()

    def export(self):
        """
        Функция экспорта приказа
        """
        # Создаем исключение на случай ошибки
        try:
            # Открываем диалоговое окно для выбора файла сохранения
            fname = QFileDialog.getSaveFileName(
                self, 'Сохранение файла', '',
                'Документ Word (*.docx);;Текстовые документы (*.txt);;Все файлы (*)')[0]
            # Список месяцев
            month_dict = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября',
                          'октября', 'ноября', 'декабря']
            # Если выбрали txt, то
            if fname[-4:] == '.txt':
                # Приводим приказывающих к формату "должность1 \t фио1 \n должность2 \t фио2 \n ..."
                guys = '\n'.join([f'{self.guysTable.item(i, 0).text()}\t{self.guysTable.item(i, 1).text()}'
                                  for i in range(self.guysTable.rowCount())])
                # Открываем файл
                with open(fname, 'w', encoding='utf8') as file:
                    # Записываем приказ
                    file.write(f'''{self.organization.text()}

ПРИКАЗ

{self.date.date().day()} {month_dict[self.date.date().month()]} {self.date.date().year()} года.\t№{self.number.text()}


{self.title.text()}


{self.reason.toPlainText()}

ПРИКАЗЫВАЮ:
{self.text.toPlainText()}

{guys}
''')
            # Если выбрали docx, то
            elif fname[-5:] == '.docx':
                # Создаем шаблон документа
                doc = DocxTemplate('template.docx')
                # Высчитываем расстояние между должностями и фио приказывающих
                space_len = lambda x: 47 - len(self.guysTable.item(x, 0).text() + self.guysTable.item(x, 1).text())
                # Вводим данные
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
                # Создаем документ
                doc.render(context)
                # Сохраняем его
                doc.save(fname)
        # Если вышла ошибка PermissionError, которая выходила у меня если файл открыт в другой программе
        except PermissionError:
            # Создаем окно
            msg_box = QMessageBox()
            # Ставим иконку
            msg_box.setIcon(QMessageBox.Warning)
            # Задаем заголовок окна
            msg_box.setWindowTitle("Внимание")
            # Задаем текст окна
            msg_box.setText("Файл открыт в другой программе.")
            # Ставим кнопку ОК
            msg_box.setStandardButtons(QMessageBox.Ok)
            # Показываем окно
            msg_box.exec_()


class AboutWindow(QWidget):
    """
    Окно с информацией о программе
    """

    def __init__(self):
        """
        Инициализация окна редактирования приказа
        """
        super().__init__()
        # Дизайн
        uic.loadUi('ui-files/aboutWidget.ui', self)
        self.setStyleSheet('.QWidget {background-image: url("images/background.jpg");}')
        self.setContentsMargins(0, 0, 0, 0)

        # Ставим мою фотографию
        photo_pixmap = QPixmap('images/myrealface.png')
        self.photo.setPixmap(photo_pixmap)
        self.photo.setFixedSize(100, 100)

        # Ставим иконку гитхаба ссылке на гитхаб
        self.githubLink.setIcon(QIcon('images/github.png'))

        # Подключаем открытие гитхаба при нажатии на ссылку
        self.githubLink.clicked.connect(lambda: webbrowser.open('https://github.com/f1rsov08/BestProject'))


# Запуск программы
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
