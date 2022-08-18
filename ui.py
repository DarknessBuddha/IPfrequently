import time

from PyQt6.QtWidgets import QDialog, QPushButton, QLineEdit, QSpinBox, QTableWidget, QTableWidgetItem, QComboBox, QLabel
from PyQt6.QtGui import QIcon
from PyQt6 import uic
import pandas
import sys
import utils
from script import run_bot
from threading import Thread

class UI(QDialog):
    def __init__(self):
        super().__init__()

        uic.loadUi(utils.resource_path('graphics.ui'), self)

        # Define Widgets
        self.file_button: QPushButton | UI = self.findChild(QPushButton, 'file_button')
        self.file_box: QLineEdit | UI = self.findChild(QLineEdit, 'file_box')
        self.sheet_box: QComboBox | UI = self.findChild(QComboBox, 'sheet_box')
        self.existing_name_box: QLineEdit | UI = self.findChild(QLineEdit, 'existing_name_box')
        self.name_box: QLineEdit | UI = self.findChild(QLineEdit, 'name_box')
        self.mac_box: QLineEdit | UI = self.findChild(QLineEdit, 'mac_box')
        self.add_range_button: QPushButton | UI = self.findChild(QPushButton, 'add_range_button')
        self.start_row_box: QSpinBox | UI = self.findChild(QSpinBox, 'start_row_box')
        self.end_row_box: QSpinBox | UI = self.findChild(QSpinBox, 'end_row_box')
        self.table: QTableWidget | UI = self.findChild(QTableWidget, 'table')
        self.start_button: QPushButton | UI = self.findChild(QPushButton, 'start_button')
        self.clear_button: QPushButton | UI = self.findChild(QPushButton, 'clear_button')
        self.console: QLabel | UI = self.findChild(QLabel, 'console')


        # Implementations
        self.file_button.clicked.connect(lambda: (
            self.sheet_box.clear(),
            self.file_box.setText(utils.get_file_name()),
            self.start_button.show(),
            self.sheet_box.addItems([name for name in pandas.read_excel(self.file_box.text(), sheet_name=None).keys()])
            if self.file_box.text() else self.start_button.hide()
        ))
        self.add_range_button.clicked.connect(lambda: (
            self.table.insertRow(self.table.rowCount()),
            self.table.setItem(self.table.rowCount() - 1, 0, QTableWidgetItem(str(self.start_row_box.value()))),
            self.table.setItem(self.table.rowCount() - 1, 1, QTableWidgetItem(str(self.end_row_box.value())))
        ))
        self.clear_button.clicked.connect(lambda: self.table.setRowCount(0))
        self.start_button.clicked.connect(lambda: Thread(target=self.run).start())



        # Settings
        self.setMinimumSize(600,444)
        # self.setFixedSize()
        self.setWindowTitle('IPfrequently')
        self.setWindowIcon(QIcon(utils.resource_path('cisco.ico')))
        self.start_button.hide()

        self.show()

    def run(self):
        self.start_button.hide()

        # countdown
        for i in range(3, 0, -1):
            self.console.setText(f'Starting in {i}')
            time.sleep(1)
        self.console.clear()

        run_bot(self.file_box.text(), self.sheet_box.currentText(),
                [[int(self.table.item(row_pos, 0).text()), int(self.table.item(row_pos, 1).text())]
                 for row_pos in range(self.table.rowCount())],
                self.existing_name_box.text(),
                self.name_box.text() or self.name_box.placeholderText(),
                self.mac_box.text() or self.mac_box.placeholderText())

        # finished
        self.console.setText('Marionette Ended')
        self.start_button.show()
        time.sleep(5)
        self.console.clear()
