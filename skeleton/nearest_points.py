# Импорт всех необходимых библиотек
import sys
import os
from datetime import date
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from backend import *


class DlgMain(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Ближайшие точки')  # Устанавливаем настройки для главного окна
        self.setFixedSize(450, 360)
        self.setFont(QFont('Century Gothic', 9))
        self.setWindowIcon(QIcon('_internal/nearest_points_ico.ico'))
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)

        self.btnFileA = QPushButton('Выбрать', self)  # Устанавливаем кнопку
        self.btnFileA.setGeometry(0, 0, 115, 35)  # Устанавливаем размеры
        self.btnFileA.move(290, 20)  # Перемещаем в нужное место
        self.btnFileA.clicked.connect(self.evt_open_file_a)  # Добавляем функцию для вызова по клику на кнопку

        self.file_a_open = QLabel('', self)  # Устанавливаем лейбл для последующего появления в нем названия файла
        self.file_a_open.setGeometry(0, 0, 115, 20)
        self.file_a_open.setStyleSheet('color: #888888;')  # Устанавливаем стиль
        self.file_a_open.setFont(QFont('Century Gothic', 7))  # Корректируем шрифт
        self.file_a_open.setAlignment(Qt.AlignCenter)  # Выравниваем по центру
        self.file_a_open.move(290, 55)

        self.label_a = QLabel('Файл с перечнем\nстанций метрополитена', self)  # Устанавливаем лейбл,
        # описывающий кнопку справа от него
        self.label_a.move(30, 15)

        self.line = QLabel('', self)  # Устанавливаем исскуственную линию
        self.line.setGeometry(0, 0, 400, 1)
        self.line.setStyleSheet('background-color: #000;')
        self.line.move(int((self.width() - self.line.width()) / 2), 80)  # Распологаем ее по центру основного окна

        self.btnFileB = QPushButton('Выбрать', self)
        self.btnFileB.setGeometry(0, 0, 115, 35)
        self.btnFileB.move(290, 100)
        self.btnFileB.clicked.connect(self.evt_open_file_b)

        self.file_b_open = QLabel('', self)
        self.file_b_open.setGeometry(0, 0, 115, 20)
        self.file_b_open.setStyleSheet('color: #888888;')
        self.file_b_open.setFont(QFont('Century Gothic', 7))
        self.file_b_open.setAlignment(Qt.AlignCenter)
        self.file_b_open.move(290, 135)

        self.label_b = QLabel('Файл с перечнем\nдостопримечательностей', self)
        self.label_b.move(30, 95)

        self.label_count = QLabel('Укажите количество \nдостопримечательностей', self)
        self.label_count.setGeometry(0, 0, 350, 50)
        self.label_count.setAlignment(Qt.AlignCenter)
        self.label_count.move(int((self.frameSize().width() - self.label_count.frameSize().width()) / 2), 170)

        self.count_input = QSpinBox(self)  # Контейнер для ввода целых чисел
        self.count_input.setGeometry(0, 0, 50, 30)
        self.count_input.move(int((self.frameSize().width() - self.count_input.frameSize().width()) / 2), 230)
        self.count_input.setMinimum(5)  # Установили минимальное значение
        self.count_input.setMaximum(40)  # Установили максимальное значение
        self.count_input.setValue(5)  # Устанавливаем значение по умолчанию
        self.count_input.setSingleStep(1)  # Устанавливаем шаг

        self.btnInput = QPushButton('Рассчитать', self)  # Устанавливаем кнопку, при клике на которую будет
        # производиться рассчет
        self.btnInput.setGeometry(0, 0, 120, 35)
        self.btnInput.setStyleSheet('background-color: #CDCDCD;')
        self.btnInput.move(int((self.frameSize().width() - self.btnInput.frameSize().width()) / 2), 300)
        self.btnInput.clicked.connect(self.evt_save)


    def evt_save(self):
        global count_points, path_a, path_b

        if 'path_a' in globals() and 'path_b' in globals():  # Проверяем, выбрал ли пользователь файлы
            count_points = int(self.count_input.text())  # Получаем числовое значение количества точек из
            # контейнера для чисел
            current_date = date.today().strftime("%d.%m.%Y")  # Запоминаем текущую дату
            create_dfs(path_a, path_b)  # оздаем фреймы данных из выбранных файлов
            if check_dfs():  # Проверяем файлы на наличие координат
                get_points_df_result(count_points)  # Передаем алгоритму количество необходимых точек,
                # там же создается пустой результирующий фрейм
                calculate()  # Производим рассчет
                # Пользователь указывает путь, по которому нужно сохранить файл. Название файла дается по умолчанию
                path_save, form = QFileDialog.getSaveFileName(self, 'Save file at', f'/nearest_points_'
                                                                    f'{current_date}.xlsx', 'EXCEL file (*.xlsx)')
                save_result(path_save)  # Сохраняем файл по выбранному пользователем пути
                res = QMessageBox.information(self, 'Информация', 'Рассчет произведен.')  # Сообщаем о
                # проведении рассчета
            else:
                res = QMessageBox.critical(self, 'Внимание!', 'Вы выбрали неподходящие файлы.')  # Сообщение о
                # неподходящих файлах
        else:
            res = QMessageBox.warning(self, 'Информация', 'Вы не выбрали файлы.')  # Сообщение об отсутствии
            # выбора файлов


    def evt_open_file_a(self):
        global path_a
        path_a, form = QFileDialog.getOpenFileName(self, 'Open file A', '/', 'EXCEL file ('
                                                                               '*.xlsx)')  # Сохраняем путь до
        # файла А
        self.file_a_open.setText(os.path.basename(path_a))  # Вписываем название файла в пустой лейбл

    def evt_open_file_b(self):  # Все также, как предыдущий файл, но для файла Б
        global path_b
        path_b, form = QFileDialog.getOpenFileName(self, 'Open file B', '/', 'EXCEL file ('
                                                                             '*.xlsx)', )
        self.file_b_open.setText(os.path.basename(path_b))



if __name__ == '__main__':
    app = QApplication(sys.argv)
    dlgMain = DlgMain()
    dlgMain.show()  # Показываем получившееся окно
    sys.exit(app.exec_())
