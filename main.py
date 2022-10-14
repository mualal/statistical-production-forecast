import sys

from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QGridLayout, QLabel


class MainWindow(QWidget):
    def __init__(
        self
    ):
        super().__init__()

        self.setWindowTitle('Статистика')
        calculate_button = QPushButton('Рассчитать')
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
