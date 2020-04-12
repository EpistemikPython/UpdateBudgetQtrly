##############################################################################################################################
# coding=utf-8
#
# startUI.py -- run the UI for the update functions
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2019-03-30'
__updated__ = '2020-04-12'

import concurrent.futures as confut
from functools import partial
from sys import argv, path
from PyQt5.QtWidgets import (QApplication, QComboBox, QVBoxLayout, QGroupBox, QDialog, QFileDialog,
                             QPushButton, QFormLayout, QDialogButtonBox, QLabel, QTextEdit, QCheckBox, QInputDialog)

path.append('/home/marksa/dev/git/Python/Gnucash/createGncTxs/')
from investment import *
from updateBudget import UPDATE_DOMAINS, SHEET_1, SHEET_2
from updateRevExps import update_rev_exps_main
from updateAssets import update_assets_main
from updateBalance import update_balance_main

# constant strings
REV_EXPS:str = 'Rev & Exps'
ASSETS:str   = 'Assets'
BALANCE:str  = 'Balance'
DOMAIN:str   = 'Domain'
DEST:str     = 'Destination'
QTRS:str     = 'Quarters'

UPDATE_FXNS = [update_rev_exps_main, update_assets_main, update_balance_main]
CHOICE_FXNS = {
    REV_EXPS : UPDATE_FXNS[0] ,
    ASSETS   : UPDATE_FXNS[1] ,
    BALANCE  : UPDATE_FXNS[2] ,
    ALL      : ALL
}


# noinspection PyCallByClass,PyTypeChecker,PyAttributeOutsideInit,PyArgumentList,PyMethodMayBeStatic
class UpdateBudgetUI(QDialog):
    """update my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file"""
    def __init__(self):
        super().__init__()
        self.title = 'Update Budget Quarterly'
        self.left = 120
        self.top  = 160
        self.width  = 600
        self.height = 800
        self.gnc_file = ''
        self.script = ''
        self.mode = ''
        self.init_ui()
        ui_lgr.info(get_current_time())

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.log_level = lg.INFO

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

    def create_group_box(self):
        self.gb_main = QGroupBox('Parameters:')
        layout = QFormLayout()

        self.cb_script = QComboBox()
        self.cb_script.addItems([x for x in CHOICE_FXNS])
        # self.cb_script.currentIndexChanged.connect(partial(self.script_change))
        layout.addRow(QLabel('Script:'), self.cb_script)
        self.script = self.cb_script.currentText()

        self.gnc_file_btn = QPushButton('Get Gnucash file')
        self.gnc_file_btn.clicked.connect(partial(self.open_file_name_dialog))
        layout.addRow(QLabel('Gnucash File:'), self.gnc_file_btn)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems([TEST,SEND])
        self.cb_mode.currentIndexChanged.connect(partial(self.mode_change))
        layout.addRow(QLabel('Mode:'), self.cb_mode)
        self.mode = self.cb_mode.currentText()

        self.cb_domain = QComboBox()
        self.cb_domain.addItems(UPDATE_DOMAINS)
        self.cb_domain.currentIndexChanged.connect(partial(self.selection_change, self.cb_domain, DOMAIN))
        layout.addRow(QLabel(DOMAIN+':'), self.cb_domain)

        self.cb_dest = QComboBox()
        self.cb_dest.currentIndexChanged.connect(partial(self.selection_change, self.cb_dest, DEST))
        layout.addRow(QLabel(DEST+':'), self.cb_dest)

        vert_box = QGroupBox('Check:')
        vert_layout = QVBoxLayout()
        self.ch_gnc = QCheckBox('Save Gnucash info to JSON file?')
        self.ch_ggl = QCheckBox('Save Google info to JSON file?')
        self.ch_rsp = QCheckBox('Save Google RESPONSE to JSON file?')

        vert_layout.addWidget(self.ch_gnc)
        vert_layout.addWidget(self.ch_ggl)
        vert_layout.addWidget(self.ch_rsp)
        vert_box.setLayout(vert_layout)
        layout.addRow(QLabel('Options'), vert_box)

        self.pb_logging = QPushButton("Change the logging level?")
        self.pb_logging.clicked.connect(self.get_log_level)
        layout.addRow(QLabel('Logging'), self.pb_logging)

        self.exe_btn = QPushButton('Go!')
        self.exe_btn.clicked.connect(partial(self.button_click))
        layout.addRow(QLabel('EXECUTE:'), self.exe_btn)

        self.gb_main.setLayout(layout)

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, 'Get Gnucash Files', '/newdata/dev/git/Python/Gnucash/app-files',
                                                   "Gnucash Files (*.gnc *.gnucash);;All Files (*)", options=options)
        if file_name:
            self.gnc_file = file_name
            gnc_file_display = file_name.split('/')[-1]
            self.gnc_file_btn.setText(gnc_file_display)

    # TODO: just have mode as 'test' or 'Sheet 1' or 'Sheet 2' and don't need separate dest widget
    def mode_change(self):
        """need the destination sheet if mode is Send"""
        new_mode = self.cb_mode.currentText()
        ui_lgr.info(F"Mode changed to '{new_mode}'.")
        if new_mode != self.mode:
            if new_mode == TEST:
                self.cb_dest.clear()
            elif new_mode == SEND:
                self.cb_dest.addItems([SHEET_1, SHEET_2])
            else:
                raise Exception(F"INVALID MODE!!?? '{new_mode}'")

            self.mode = new_mode

    def get_log_level(self):
        num, ok = QInputDialog.getInt(self, "Logging Level", "Enter a value (0-100)", value=self.log_level, min=0, max=100)
        if ok:
            self.log_level = num
            ui_lgr.info(F"logging level changed to {num}.")

    def selection_change(self, cb:QComboBox, label:str):
        """info printing only"""
        ui_lgr.info(F"ComboBox '{label}' selection changed to '{cb.currentText()}'.")

    def thread_update(self, thread_fxn:object, p_params:list):
        ui_lgr.info(F"starting thread: {str(thread_fxn)}")
        if callable(thread_fxn):
            response = thread_fxn(p_params)
        else:
            msg = F"thread fxn {str(thread_fxn)} NOT callable?!"
            ui_lgr.warning(msg)
            raise Exception(msg)
        ui_lgr.info(F"finished thread: {str(thread_fxn)}\n")
        return response

    def button_click(self):
        """assemble the necessary parameters"""
        ui_lgr.info(F"Clicked '{self.exe_btn.text()}'.")
        exe = self.cb_script.currentText()
        ui_lgr.info(F"Script is '{exe}'.")

        if not self.gnc_file:
            self.response_box.append('>>> MUST select a Gnucash File!')
            return

        cl_params = ['-g' + self.gnc_file]

        # if sending, adjust to the Sheet selected
        if self.cb_mode.currentText() == SEND:
            send_mode = self.cb_dest.currentText()
            # save Google response
            if self.ch_rsp.isChecked(): cl_params.append('--resp_save')
        else:
            send_mode = TEST
        cl_params.append('-m' + send_mode)

        cl_params.append('-t' + self.cb_domain.currentText())
        cl_params.append('-l' + str(self.log_level))

        if self.ch_ggl.isChecked() : cl_params.append('--ggl_save')
        if self.ch_gnc.isChecked() : cl_params.append('--gnc_save')

        ui_lgr.info(str(cl_params))

        main_fxn = CHOICE_FXNS[exe]
        if callable(main_fxn):
            ui_lgr.info('Calling main function...')
            response = main_fxn(cl_params)
            reply = {'response': response}
        elif main_fxn == ALL:
            ui_lgr.info('main_fxn == ALL')
            # use 'with' to ensure threads are cleaned up properly
            with confut.ThreadPoolExecutor(max_workers = len(UPDATE_FXNS)) as executor:
                # send each script to a separate thread
                future_to_update = {executor.submit(self.thread_update, upfxn, cl_params):upfxn for upfxn in UPDATE_FXNS}
                for future in confut.as_completed(future_to_update):
                    updater = future_to_update[future]
                    try:
                        data = future.result()
                    except Exception as bcae:
                        msg = repr(bcae)
                        ui_lgr.warning(F"{updater} generated exception: {msg}")
                        raise bcae
                    else:
                        reply = data
                        ui_lgr.info(F"Updater '{updater}' has finished.")
        else:
            msg = F"Problem with main??!! '{main_fxn}'"
            ui_lgr.error(msg)
            reply = {'msg': msg, 'log': saved_log_info}

        self.response_box.append(json.dumps(reply, indent=4))

# END class UpdateBudgetUI


def ui_main():
    app = QApplication(argv)
    dialog = UpdateBudgetUI()
    dialog.show()
    app.exec_()


if __name__ == '__main__':
    ui_lgr = get_logger(UpdateBudgetUI.__name__)
    # ui_lgr.setLevel(13)
    ui_main()
    finish_logging(UpdateBudgetUI.__name__)
    exit()
