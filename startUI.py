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
from updateCommon import *
from updateRevExps import update_rev_exps_main
from updateAssets import update_assets_main
from updateBalance import update_balance_main


# constant strings
REV_EXPS  = 'Rev & Exps'
ASSETS    = 'Assets'
BALANCE   = 'Balance'
TEST      = 'test'
SEND      = 'send'
GNC_FILES = 'Script'
QRTRS     = 'Quarters'

PARAMS = {
    GNC_FILES : ['reader', 'runner', 'HouseHold'] ,
    REV_EXPS  : ['2019', '2018', '2017', '2016', '2015', '2014', '2013', '2012'] ,
    ASSETS    : ['2011', '2010', '2009', '2008'] ,
    BALANCE   : ['today', 'allyears'] ,
    QRTRS     : ['0', '1', '2', '3', '4']
}


# noinspection PyAttributeOutsideInit,PyUnresolvedReferences
class UpdateBudgetQtrly(QDialog):

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

        self.response = QTextEdit()
        self.response.setReadOnly(True)
        self.response.setText('Hello there!')

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Close)
        button_box.accepted.connect(self.accept)
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
        self.cb_script.currentIndexChanged.connect(partial(self.script_change))
        layout.addRow(QLabel("Script:"), self.cb_script)
        self.script = self.cb_script.currentText()

        self.cb_gnc_file = QComboBox()
        self.cb_gnc_file.addItems(PARAMS[GNC_FILES])
        self.cb_gnc_file.currentIndexChanged.connect(partial(file_change, self.cb_gnc_file))
        layout.addRow(QLabel("Gnucash File:"), self.cb_gnc_file)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems([TEST, SEND])
        self.cb_mode.currentIndexChanged.connect(partial(self.mode_change))
        layout.addRow(QLabel("Mode:"), self.cb_mode)
        self.mode = self.cb_mode.currentText()

        self.cb_domain = QComboBox()
        self.cb_domain.addItems(PARAMS[REV_EXPS])
        self.cb_domain.currentIndexChanged.connect(partial(domain_change, self.cb_domain))
        layout.addRow(QLabel("Domain:"), self.cb_domain)

        self.cb_qtr = QComboBox()
        self.cb_qtr.addItems(PARAMS[QRTRS])
        self.cb_qtr.currentIndexChanged.connect(partial(quarter_change, self.cb_qtr))
        layout.addRow(QLabel("Quarter:"), self.cb_qtr)

        self.cb_dest = QComboBox()
        self.cb_dest.currentIndexChanged.connect(partial(dest_change, self.cb_dest))
        layout.addRow(QLabel("Destination:"), self.cb_dest)

        self.exe_btn = QPushButton('Go!')
        self.exe_btn.clicked.connect(partial(button_click, self))
        layout.addRow(QLabel("Execute:"), self.exe_btn)

        self.formGroupBox.setLayout(layout)

    def script_change(self):
        new_script = self.cb_script.currentText()
        print_info("Script changed to '{}'.".format(new_script), MAGENTA)
        if new_script != self.script:
            if new_script == REV_EXPS:
                # adjust Domain
                self.cb_domain.clear()
                self.cb_domain.addItems(PARAMS[REV_EXPS])
                # adjust Quarter if necessary
                if self.script == BALANCE:
                    self.cb_qtr = PARAMS[QRTRS]
            elif new_script == ASSETS:
                # adjust Domain
                if self.script == REV_EXPS:
                    self.cb_domain.addItems(PARAMS[ASSETS])
                else: # current script is BALANCE
                    self.cb_domain.clear()
                    self.cb_domain.addItems(PARAMS[REV_EXPS] + PARAMS[ASSETS])
                    # adjust Quarter
                    self.cb_qtr = PARAMS[QRTRS]
            elif new_script == BALANCE:
                # adjust Domain
                self.cb_domain.clear()
                self.cb_domain.addItems(PARAMS[BALANCE] + PARAMS[REV_EXPS] + PARAMS[ASSETS])
                # adjust Quarter
                self.cb_qtr.clear()
            else:
                raise Exception("INVALID SCRIPT!!?? '{}'".format(new_script))

            self.script = new_script

    def mode_change(self):
        new_mode = self.cb_mode.currentText()
        print_info("Mode changed to '{}'.".format(new_mode), CYAN)
        if new_mode != self.mode:
            if new_mode == TEST:
                self.cb_dest.clear()
            elif new_mode == SEND:
                self.cb_dest.addItems(['Sheet 1', 'Sheet 2'])
            else:
                raise Exception("INVALID MODE!!?? '{}'".format(new_mode))

            self.mode = new_mode


def domain_change(cb):
    print("Domain changed to '{}'.".format(cb.currentText()))


def file_change(cb):
    print("File changed to '{}'.".format(cb.currentText()))


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
    dialog = UpdateBudgetQtrly()
    dialog.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    ui_main()

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     dialog = Dialog()
#     sys.exit(dialog.exec_())
