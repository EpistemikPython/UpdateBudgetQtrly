##############################################################################################################################
# coding=utf-8
#
# startUI.py -- run the UI for the update functions
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__  = 3.9
__gnucash_version__ = 3.8
__created__ = '2019-03-30'
__updated__ = '2020-01-11'

from sys import argv, path
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from PyQt5.QtWidgets import ( QApplication, QComboBox, QVBoxLayout, QGroupBox, QDialog, QFileDialog,
                              QPushButton, QFormLayout, QDialogButtonBox, QLabel, QTextEdit, QCheckBox )
from functools import partial
from updateRevExps import update_rev_exps_main
from updateAssets import update_assets_main
from updateBalance import update_balance_main
from investment import *


# constant strings
REV_EXPS:str = 'Rev & Exps'
ASSETS:str   = 'Assets'
BALANCE:str  = 'Balance'
DOMAIN:str   = 'Domain'
DEST:str     = 'Destination'
QTRS:str     = 'Quarters'
SHEET_1:str  = 'Sheet 1'
SHEET_2:str  = 'Sheet 2'

PARAMS:dict = {
    REV_EXPS  : ['2019', '2018', '2017', '2016', '2015', '2014', '2013', '2012'] ,
    ASSETS    : ['2011', '2010', '2009', '2008'] ,
    BALANCE   : ['today', 'allyears'] ,
    QTRS      : ['ALL', '#1', '#2', '#3', '#4']
}

MAIN_FXNS:dict = {
    REV_EXPS : update_rev_exps_main ,
    ASSETS   : update_assets_main   ,
    BALANCE  : update_balance_main
}


# noinspection PyCallByClass,PyTypeChecker,PyArgumentList
class UpdateBudgetQtrly(QDialog):
    """update my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file"""
    def __init__(self):
        super().__init__()
        self.title = 'Update Budget Quarterly'
        self.left = 680
        self.top = 160
        self.width = 600
        self.height = 800
        self.gnc_file = ''
        self.script = ''
        self.mode = ''
        self.log = SattoLog(do_printing=True)
        self.log.print_info("{}".format(self.title), GREEN)
        self.init_ui()

    # noinspection PyAttributeOutsideInit,PyArgumentList
    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.create_group_box()

        self.response_box = QTextEdit()
        self.response_box.setReadOnly(True)
        self.response_box.acceptRichText()
        self.response_box.setText('Hello there!')

        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.gb_main)
        layout.addWidget(self.response_box)
        layout.addWidget(button_box)

        self.setLayout(layout)
        self.show()

    # noinspection PyAttributeOutsideInit,PyUnresolvedReferences
    def create_group_box(self):
        self.gb_main = QGroupBox("Parameters:")
        layout = QFormLayout()

        self.cb_script = QComboBox()
        self.cb_script.addItems([x for x in MAIN_FXNS])
        self.cb_script.currentIndexChanged.connect(partial(self.script_change))
        layout.addRow(QLabel("Script:"), self.cb_script)
        self.script = self.cb_script.currentText()

        self.gnc_file_btn = QPushButton('Get Gnucash file')
        self.gnc_file_btn.clicked.connect(partial(self.open_file_name_dialog))
        layout.addRow(QLabel("Gnucash File:"), self.gnc_file_btn)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems([TEST,SEND])
        self.cb_mode.currentIndexChanged.connect(partial(self.mode_change))
        layout.addRow(QLabel("Mode:"), self.cb_mode)
        self.mode = self.cb_mode.currentText()

        self.cb_domain = QComboBox()
        self.cb_domain.addItems(PARAMS[REV_EXPS])
        self.cb_domain.currentIndexChanged.connect(partial(self.selection_change, self.cb_domain, DOMAIN))
        layout.addRow(QLabel(DOMAIN+':'), self.cb_domain)

        self.cb_qtr = QComboBox()
        self.cb_qtr.addItems(PARAMS[QTRS])
        self.cb_qtr.currentIndexChanged.connect(partial(self.selection_change, self.cb_qtr, QTR))
        layout.addRow(QLabel(QTR+':'), self.cb_qtr)

        self.cb_dest = QComboBox()
        self.cb_dest.currentIndexChanged.connect(partial(self.selection_change, self.cb_dest, DEST))
        layout.addRow(QLabel(DEST+':'), self.cb_dest)

        vert_box = QGroupBox("Check:")
        vert_layout = QVBoxLayout()
        self.ch_gnc = QCheckBox("Save Gnucash info to JSON file?")
        self.ch_ggl = QCheckBox("Save Google info to JSON file?")
        self.ch_debug = QCheckBox("Print DEBUG info?")
        vert_layout.addWidget(self.ch_gnc)
        vert_layout.addWidget(self.ch_ggl)
        vert_layout.addWidget(self.ch_debug)
        vert_box.setLayout(vert_layout)
        layout.addRow(QLabel("Options"), vert_box)

        self.exe_btn = QPushButton('Go!')
        self.exe_btn.clicked.connect(partial(self.button_click))
        layout.addRow(QLabel("Execute:"), self.exe_btn)

        self.gb_main.setLayout(layout)

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Get Gnucash Files", "",
                                                   "Gnucash Files (*.gnc);;All Files (*)", options=options)
        if file_name:
            self.gnc_file = file_name
            gnc_file_display = file_name.split('/')[-1]
            self.gnc_file_btn.setText(gnc_file_display)

    def script_change(self):
        """must adjust domain and possibly quarter"""
        new_script = self.cb_script.currentText()
        self.log.print_info("Script changed to: {}".format(new_script), MAGENTA)
        if new_script != self.script:
            initial_domain = self.cb_domain.currentText()
            self.log.print_info("Start with domain = {}".format(initial_domain), BROWN)
            if new_script == REV_EXPS:
                self.cb_domain.clear()
                self.cb_domain.addItems(PARAMS[REV_EXPS])
                if self.script == BALANCE:
                    self.cb_qtr.clear()
                    self.cb_qtr.addItems(PARAMS[QTRS])
            else:
                if self.script == REV_EXPS:
                    self.cb_domain.addItems(PARAMS[ASSETS])
                if new_script == ASSETS:
                    if self.script == BALANCE:
                        self.cb_domain.clear()
                        self.cb_domain.addItems(PARAMS[REV_EXPS] + PARAMS[ASSETS])
                        self.cb_qtr.clear()
                        self.cb_qtr.addItems(PARAMS[QTRS])
                elif new_script == BALANCE:
                    self.cb_domain.addItems(PARAMS[BALANCE])
                    self.cb_qtr.clear()
                    self.cb_qtr.addItem('NOT NEEDED')
                else:
                    raise Exception("INVALID SCRIPT!!?? '{}'".format(new_script))

            # does not seem to be any defined function to return list of current items
            if self.cb_domain.currentText() != initial_domain \
                    and initial_domain in [self.cb_domain.itemText(i) for i in range(self.cb_domain.count())]:
                self.cb_domain.setCurrentText(initial_domain)

            self.log.print_info("Finish with domain = {}".format(self.cb_domain.currentText()), BROWN)
            self.script = new_script

    def mode_change(self):
        """need the destination sheet if mode is Send"""
        new_mode = self.cb_mode.currentText()
        self.log.print_info("Mode changed to '{}'.".format(new_mode), CYAN)
        if new_mode != self.mode:
            if new_mode == TEST:
                self.cb_dest.clear()
            elif new_mode == SEND:
                self.cb_dest.addItems([SHEET_1, SHEET_2])
            else:
                raise Exception("INVALID MODE!!?? '{}'".format(new_mode))

            self.mode = new_mode

    def button_click(self):
        """assemble the necessary parameters"""
        self.log.print_info("Clicked '{}'.".format(self.exe_btn.text()))
        exe = self.cb_script.currentText()
        self.log.print_info("Script is '{}'.".format(exe))

        if self.gnc_file == '':
            self.response_box.setText('>>> MUST select a Gnucash File!')
            return

        # if sending, adjust the mode string to match the Sheet selected
        send_mode = self.cb_mode.currentText()
        if send_mode == SEND:
            if self.cb_dest.currentText() == SHEET_1:
                send_mode += '1'
            elif self.cb_dest.currentText() == SHEET_2:
                send_mode += '2'

        cl_params = ['-g' + self.gnc_file, '-m' + send_mode]

        quarter = self.cb_qtr.currentText().replace('#','')
        domain_key = '-p'
        # BALANCE has slightly different parameters than REV_EXPS and ASSETS
        if exe != BALANCE:
            domain_key = '-y'
            if quarter != 'ALL' : cl_params.append('-q' + quarter)
            if self.ch_gnc.isChecked() : cl_params.append('--gnc_save')

        domain = self.cb_domain.currentText()
        cl_params.append(domain_key + domain)

        if self.ch_ggl.isChecked() : cl_params.append('--ggl_save')
        if self.ch_debug.isChecked() : cl_params.append('--debug')

        self.log.print_info(cl_params, GREEN)

        main_fxn = MAIN_FXNS[exe]
        if callable(main_fxn):
            reply = main_fxn(cl_params)
        else:
            msg = "Problem with main??!! '{}'".format(main_fxn)
            self.log.print_error(msg)
            reply = msg

        self.response_box.append(json.dumps(reply, indent=4))

    def selection_change(self, cb:QComboBox, label:str):
        """info printing only"""
        self.log.print_info("ComboBox '{}' selection changed to '{}'.".format(label, cb.currentText()), BLUE)


def ui_main():
    app = QApplication(argv)
    dialog = UpdateBudgetQtrly()
    dialog.show()
    exit(app.exec_())


if __name__ == '__main__':
    ui_main()

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     dialog = Dialog()
#     sys.exit(dialog.exec_())
