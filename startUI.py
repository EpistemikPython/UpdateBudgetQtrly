import sys
from PyQt5.QtWidgets import ( QApplication, QWidget, QComboBox, QVBoxLayout, QGroupBox, QGridLayout, QDialog,
                              QPushButton, QFormLayout, QDialogButtonBox, QLabel, QLineEdit, QSpinBox )
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
        self.height = 400
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.create_grid_layout()

        self.createFormGroupBox()

        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.formGroupBox)
        layout.addWidget(self.group_box)
        layout.addWidget(buttonBox)
        self.setLayout(layout)

        self.setLayout(layout)
        self.show()

    def create_grid_layout(self):
        self.group_box = QGroupBox('Selections:')
        layout = QGridLayout()
        # layout.setColumnStretch(1, 4)
        # layout.setColumnStretch(2, 4)

        self.cb_script = QComboBox()
        self.cb_script.addItems([REV_EXPS, ASSETS, BALANCE])
        self.cb_script.currentIndexChanged.connect(partial(script_change, self.cb_script))
        layout.addWidget(self.cb_script, 0, 0)

        self.cb_gnc_file = QComboBox()
        self.cb_gnc_file.addItems(['reader', 'runner', 'HouseHold'])
        self.cb_gnc_file.currentIndexChanged.connect(partial(file_change, self.cb_gnc_file))
        layout.addWidget(self.cb_gnc_file, 0, 1)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems(['test', 'send'])
        self.cb_mode.currentIndexChanged.connect(partial(mode_change, self.cb_mode))
        layout.addWidget(self.cb_mode, 1, 0)

        layout.addWidget(QLabel('BLANK'), 1, 1)

        self.cb_domain = QComboBox()
        self.cb_domain.addItems(DOMAINS[REV_EXPS])
        self.cb_domain.currentIndexChanged.connect(partial(domain_change, self.cb_domain))
        layout.addWidget(self.cb_domain, 2, 0)

        self.cb_qtr = QComboBox()
        self.cb_qtr.addItems(['0', '1', '2', '3', '4'])
        self.cb_qtr.currentIndexChanged.connect(partial(quarter_change, self.cb_qtr))
        layout.addWidget(self.cb_qtr, 2, 1)

        self.cb_dest = QComboBox()
        self.cb_dest.addItems(['Sheet 1', 'Sheet 2'])
        self.cb_dest.currentIndexChanged.connect(partial(dest_change, self.cb_dest))
        layout.addWidget(self.cb_dest, 3, 0)

        self.btn = QPushButton('Execute')
        layout.addWidget(self.btn, 3, 1)
        self.btn.clicked.connect(partial(button_click, self))

        self.group_box.setLayout(layout)

    NumGridRows = 3
    NumButtons = 4

    def createFormGroupBox(self):
        self.formGroupBox = QGroupBox("Form layout")
        layout = QFormLayout()

        self.cb_year = QComboBox()
        self.cb_year.addItems(DOMAINS[ASSETS])
        self.cb_year.currentIndexChanged.connect(partial(year_change, self.cb_year))

        layout.addRow(QLabel("Name:"), QLineEdit())
        layout.addRow(QLabel("Year:"), self.cb_year)
        layout.addRow(QLabel("Age:"), QSpinBox())
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
    print("Script is '{}'.".format(obj.cb_script.currentText()))
    print("Clicked '{}'.".format(obj.btn.text()))


def ui_main():
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    ui_main()


# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     dialog = Dialog()
#     sys.exit(dialog.exec_())
