##############################################################################################################################
# coding=utf-8
#
# pyside6-UI.py
#   -- access the Google Sheet update functions using a UI built with the PySide6 Qt library
#
# Copyright (c) 2025 Mark Sattolo <epistemik@gmail.com>

__author_name__    = "Mark Sattolo"
__author_email__   = "epistemik@gmail.com"
__python_version__ = "3.10+"
__created__ = "2024-07-01"
__updated__ = "2025-08-01"

from sys import path
from PySide6.QtWidgets import (QApplication, QComboBox, QVBoxLayout, QGroupBox, QDialog, QFileDialog, QLabel, QCheckBox,
                               QPushButton, QFormLayout, QDialogButtonBox, QTextEdit, QInputDialog, QMessageBox)
from functools import partial
path.append("/home/marksa/git/Python/utils")
from updateBudget import *
from updateRevExps import update_rev_exps_main
from updateAssets import update_assets_main
from updateBalance import update_balance_main

TIMEFRAME:str = "Time Frame"
UPDATE_DOMAINS = [CURRENT_YRS, RECENT_YRS, MID_YRS, EARLY_YRS, ALL_YEARS] + [year for year in UPDATE_YEARS]
UPDATE_FXNS = [update_rev_exps_main, update_assets_main, update_balance_main]
FXNS_TABLE = {
    BAL+' & '+ASSET+'s' : UPDATE_FXNS[1:] ,
    ALL                 : UPDATE_FXNS ,
    BAL                 : UPDATE_FXNS[2] ,
    ASSET+'s'           : UPDATE_FXNS[1] ,
    "Rev & Exps"        : UPDATE_FXNS[0]
}
UI_DEFAULT_LOG_LEVEL:int = logging.INFO


# noinspection PyAttributeOutsideInit
class UpdateBudgetUI(QDialog):
    """UI for updating my 'Budget Quarterly' Google sheet with information from a Gnucash file."""
    def __init__(self):
        super().__init__()
        self.title = "Update Budget UI"
        self.left = 20
        self.top  = 120
        self.width  = 620
        self.height = 800
        self.gnc_file = ""

        self._lgr = log_control.get_logger()
        self._lgr.log(UI_DEFAULT_LOG_LEVEL, f"{self.title} runtime = {get_current_time()}" )

        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.selected_loglevel = UI_DEFAULT_LOG_LEVEL
        self.create_group_box()

        self.response_box = QTextEdit()
        self.response_box.setReadOnly(True)
        self.response_box.acceptRichText()
        self.response_box.setText("Hello there!")

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.gb_main)
        layout.addWidget(self.response_box)
        layout.addWidget(button_box)
        self.setLayout(layout)

    def create_group_box(self):
        self.gb_main = QGroupBox("Parameters:")
        layout = QFormLayout()

        self.cb_script = QComboBox()
        self.cb_script.addItems( list(FXNS_TABLE.keys()) )
        layout.addRow(QLabel("Script:"), self.cb_script)

        self.gnc_file_btn = QPushButton("Get Gnucash file")
        self.gnc_file_btn.clicked.connect(self.open_file_name_dialog)
        layout.addRow(QLabel("Gnucash File:"), self.gnc_file_btn)

        self.cb_target = QComboBox()
        self.cb_target.addItems([TEST, SHEET_1, SHEET_2])
        self.cb_target.currentIndexChanged.connect(partial(self.selection_change, self.cb_target, TARGET))
        layout.addRow(QLabel(TARGET+':'), self.cb_target)

        self.cb_domain = QComboBox()
        self.cb_domain.addItems(UPDATE_DOMAINS)
        self.cb_domain.currentIndexChanged.connect(partial(self.selection_change, self.cb_domain, TIMEFRAME))
        layout.addRow(QLabel(TIMEFRAME+':'), self.cb_domain)

        vert_box = QGroupBox("Check:")
        vert_layout = QVBoxLayout()
        self.ch_gnc = QCheckBox("Save Gnucash info to JSON file?")
        self.ch_ggl = QCheckBox("Save Google info to JSON file?")
        self.ch_rsp = QCheckBox("Save Google RESPONSE to JSON file?")

        vert_layout.addWidget(self.ch_gnc)
        vert_layout.addWidget(self.ch_ggl)
        vert_layout.addWidget(self.ch_rsp)
        vert_box.setLayout(vert_layout)
        layout.addRow(QLabel("Options"), vert_box)

        self.pb_logging = QPushButton("Change the logging level?")
        self.pb_logging.clicked.connect(self.get_log_level)
        layout.addRow(QLabel("Logging"), self.pb_logging)

        self.exec_btn = QPushButton("Go!")
        self.exec_btn.setStyleSheet("QPushButton {font-weight: bold; color: red; background-color: yellow;}")
        self.exec_btn.clicked.connect(self.button_click)
        layout.addRow(QLabel("EXECUTE:"), self.exec_btn)

        self.gb_main.setLayout(layout)

    def open_file_name_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName( self, "Get Gnucash Files", osp.join(BASE_GNUCASH_FOLDER, "bak-files" + osp.sep),
                                                    "Gnucash Files (*.gnc *.gnucash);;All Files (*)",
                                                    options = QFileDialog.Option.DontUseNativeDialog )
        if file_name:
            self.gnc_file = file_name
            gnc_file_display = get_filename(file_name)
            self.gnc_file_btn.setText(gnc_file_display)

    def get_log_level(self):
        num, ok = QInputDialog.getInt(self, "Logging Level", "Enter a value (0-60)",
                                      value=self.selected_loglevel, minValue=0, maxValue=60)
        if ok:
            self.selected_loglevel = num
            self._lgr.debug(f"logging level changed to {num}.")

    # ? 'partial' always passes the index of the chosen label as an extra param...!
    def selection_change(self, cb:QComboBox, label:str, indx:int):
        self._lgr.debug(f"ComboBox '{label}' selection changed to: {cb.currentText()} [{indx}].")

    def button_click(self):
        """Assemble the necessary parameters and call each selected update choice separately."""
        self._lgr.info(f"Clicked '{self.exec_btn.text()}'.")
        if not self.gnc_file:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setText("MUST select a Gnucash file!")
            msg_box.exec()
            return

        cl_params = [ '-g' + self.gnc_file, '-m' + self.cb_target.currentText(), '-t' + self.cb_domain.currentText(),
                      '-l' + str(self.selected_loglevel) ]
        if self.ch_ggl.isChecked(): cl_params.append("--ggl_save")
        if self.ch_gnc.isChecked(): cl_params.append("--gnc_save")
        if self.ch_rsp.isChecked(): cl_params.append("--resp_save")
        self._lgr.info(f"parameters = {repr(cl_params)}")

        exe = self.cb_script.currentText()
        self._lgr.info(f"updates to run = '{exe}'")
        main_run = FXNS_TABLE[exe]
        if callable(main_run):
            self._lgr.info(f"Calling {exe}...")
            response = main_run(cl_params)
            self.response_box.append(json.dumps({f"{main_run}\n":response}, indent = 4))
        elif isinstance(main_run, list):
            try:
                for bc_exec in main_run:
                    self._lgr.info(f"Calling '{repr(bc_exec)}' ...")
                    response = bc_exec(cl_params)
                    self.response_box.append(json.dumps({f"{bc_exec}\n":response}, indent = 4))
            except Exception as bcex:
                self.response_box.append(f"EXCEPTION:\n{repr(bcex)}")
                raise bcex
        else:
            msg = f"Problem with functions??!! '{exe}'"
            self._lgr.error(msg)
            self.response_box.append(msg)
# END class UpdateBudgetUI


if __name__ == "__main__":
    log_control = MhsLogger( UpdateBudgetUI.__name__, con_level = UI_DEFAULT_LOG_LEVEL, suffix = DEFAULT_LOG_SUFFIX )
    dialog = None
    app = None
    code = 0
    try:
        app = QApplication(argv)
        dialog = UpdateBudgetUI()
        dialog.show()
        app.exec()
    except KeyboardInterrupt as mki:
        log_control.exception(mki)
        code = 13
    except ValueError as mve:
        log_control.error(mve)
        code = 27
    except Exception as mex:
        log_control.exception(mex)
        code = 66
    finally:
        if dialog:
            dialog.close()
        if app:
            app.exit(code)
    exit(code)
