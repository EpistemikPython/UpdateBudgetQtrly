#
# startUI.py -- run the UI for the update functions
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-21
# @updated 2019-04-21

import sys
from PyQt5.QtWidgets import ( QApplication, QComboBox, QVBoxLayout, QGroupBox, QDialog,
                              QPushButton, QFormLayout, QDialogButtonBox, QLabel, QTextEdit )
from functools import partial

REV_EXPS = 'Rev & Exps'
ASSETS   = 'Assets'
BALANCE  = 'Balance'

DOMAINS = {
    REV_EXPS: ['2019', '2018', '2017', '2016', '2015', '2014', '2013', '2012'] ,
    ASSETS  : ['2011', '2010', '2009', '2008'] ,
    BALANCE : ['today', 'allyears']
}


# noinspection PyAttributeOutsideInit,PyUnresolvedReferences
class App(QDialog):

    def __init__(self):
        super().__init__()
        self.title = 'Update Budget Quarterly'
        self.left = 600
        self.top = 300
        self.width = 400
        self.height = 600
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.create_group_box()

        exe_btn = QPushButton('Execute')
        exe_btn.clicked.connect(partial(button_click, self))

        self.response = QTextEdit()
        self.response.setReadOnly(True)
        self.response.setText('Hello there!')

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Close)
        button_box.accepted.connect(partial(button_click, self))
        button_box.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.formGroupBox)
        layout.addWidget(self.response)
        layout.addWidget(button_box)

        self.setLayout(layout)
        self.show()

    def create_group_box(self):

        self.formGroupBox = QGroupBox("Parameters:")
        layout = QFormLayout()

        self.cb_script = QComboBox()
        self.cb_script.addItems([REV_EXPS, ASSETS, BALANCE])
        self.cb_script.currentIndexChanged.connect(partial(script_change, self.cb_script))
        layout.addRow(QLabel("Script:"), self.cb_script)

        self.cb_gnc_file = QComboBox()
        self.cb_gnc_file.addItems(['reader', 'runner', 'HouseHold'])
        self.cb_gnc_file.currentIndexChanged.connect(partial(file_change, self.cb_gnc_file))
        layout.addRow(QLabel("Gnucash File:"), self.cb_gnc_file)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems(['test', 'send'])
        self.cb_mode.currentIndexChanged.connect(partial(mode_change, self.cb_mode))
        layout.addRow(QLabel("Mode:"), self.cb_mode)

        self.cb_domain = QComboBox()
        self.cb_domain.addItems(DOMAINS[REV_EXPS])
        self.cb_domain.currentIndexChanged.connect(partial(domain_change, self.cb_domain))
        layout.addRow(QLabel("Domain:"), self.cb_domain)

        self.cb_qtr = QComboBox()
        self.cb_qtr.addItems(['0', '1', '2', '3', '4'])
        self.cb_qtr.currentIndexChanged.connect(partial(quarter_change, self.cb_qtr))
        layout.addRow(QLabel("Quarter:"), self.cb_qtr)

        self.cb_dest = QComboBox()
        self.cb_dest.addItems(['Sheet 1', 'Sheet 2'])
        self.cb_dest.currentIndexChanged.connect(partial(dest_change, self.cb_dest))
        layout.addRow(QLabel("Destination:"), self.cb_dest)

        self.formGroupBox.setLayout(layout)


def selection_change(cb):
    print("Selection changed to '{}'.".format(cb.currentText()))


def script_change(cb):
    print("Script changed to '{}'.".format(cb.currentText()))


def domain_change(cb):
    print("Domain changed to '{}'.".format(cb.currentText()))


def file_change(cb):
    print("File changed to '{}'.".format(cb.currentText()))


def mode_change(cb):
    print("Mode changed to '{}'.".format(cb.currentText()))


def year_change(cb):
    print("Year changed to '{}'.".format(cb.currentText()))


def quarter_change(cb):
    print("Quarter changed to '{}'.".format(cb.currentText()))


def dest_change(cb):
    print("Destination changed to '{}'.".format(cb.currentText()))


def button_click(obj):
    print("Clicked 'Execute'.")
    print("Script is '{}'.".format(obj.cb_script.currentText()))


def ui_main():
    app = QApplication(sys.argv)
    exe = App()
    exe.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    ui_main()

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     dialog = Dialog()
#     sys.exit(dialog.exec_())
