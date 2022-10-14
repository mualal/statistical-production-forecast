import sys
import os
import pandas as pd
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QGridLayout, QLabel, QFileDialog
from helpful_tools import history_preprocessing


def choose_file_with_monthly_operating_report():
    file_path, check = QFileDialog.getOpenFileName(
        parent=None,
        caption='Выберите файл с МЭР',
        directory=os.path.dirname(os.getcwd()).replace(os.sep, '/'),
        filter='Excel files (*.xlsx)'
    )
    if check:
        df_initial = pd.read_excel(file_path, sheet_name='МЭР')
        df_initial = history_preprocessing(df_initial, 365)
        print(df_initial)
    

class MainWindow(QWidget):
    def __init__(
        self
    ):
        super().__init__()

        self.setWindowTitle('Статистика')

        calculate_button = QPushButton('Рассчитать')
        calculate_button.clicked.connect(choose_file_with_monthly_operating_report)

        download_button = QPushButton('Загрузить')

        tool_label = QLabel(
            'Инструмент для прогнозирования показателей \nбазовой добычи нефти'
            ' и обводнённости на основе \nмесячного эксплуатационного рапорта (МЭР)'
        )

        self.setFixedSize(QSize(600, 200))

        grid_box = QGridLayout()
        grid_box.addWidget(calculate_button, 0, 0)
        grid_box.addWidget(download_button, 0, 1)
        grid_box.addWidget(tool_label, 1, 0, 1, 2)

        self.setLayout(grid_box)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()
