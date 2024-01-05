##############################################################################################################################
# coding=utf-8
#
# startUI.py -- run the UI for the update functions
#
# Copyright (c) 2024 Mark Sattolo <epistemik@gmail.com>

__author_name__    = "Mark Sattolo"
__author_email__   = "epistemik@gmail.com"
__python_version__ = "3.6+"
__created__ = "2019-03-30"
__updated__ = "2024-01-05"

from sys import path
from PyQt5.QtWidgets import (QApplication, QComboBox, QVBoxLayout, QGroupBox, QDialog, QFileDialog,
                             QPushButton, QFormLayout, QDialogButtonBox, QLabel, QTextEdit, QCheckBox, QInputDialog)
from functools import partial
import concurrent.futures as confut
path.append("/home/marksa/git/Python/utils")
from updateBudget import *
from updateRevExps import update_rev_exps_main
from updateAssets import update_assets_main
from updateBalance import update_balance_main

TIMEFRAME:str = "Time Frame"
UPDATE_DOMAINS = [CURRENT_YRS, RECENT_YRS, MID_YRS, EARLY_YRS, ALL_YEARS] + [year for year in UPDATE_YEARS]
UPDATE_FXNS = [update_rev_exps_main, update_assets_main, update_balance_main]
CHOICE_FXNS = {
    BAL+' & '+ASSET+'s' : UPDATE_FXNS[1:] ,
    ALL                 : UPDATE_FXNS ,
    BAL                 : UPDATE_FXNS[2] ,
    ASSET+'s'           : UPDATE_FXNS[1] ,
    "Rev & Exps"        : UPDATE_FXNS[0]
}
UI_DEFAULT_LOG_LEVEL = logging.INFO


# noinspection PyAttributeOutsideInit
class UpdateBudgetUI(QDialog):
    """UI for updating my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file."""
    def __init__(self):
        super().__init__()
        self.title = "Update Budget UI"
        self.left = 120
        self.top  = 160
        self.width  = 600
        self.height = 800
        self.gnc_file = ""
        self.init_ui()
        lg_ctrl.show( F"{self.title} runtime = {get_current_time()}" )

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.log_level:int = UI_DEFAULT_LOG_LEVEL
        self.create_group_box()

        self.response_box = QTextEdit()
        self.response_box.setReadOnly(True)
        self.response_box.acceptRichText()
        self.response_box.setText("Hello there!")

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
        self.gb_main = QGroupBox("Parameters:")
        layout = QFormLayout()

        self.cb_script = QComboBox()
        self.cb_script.addItems([x for x in CHOICE_FXNS.keys()])
        # self.cb_script.currentIndexChanged.connect(partial(self.script_change))
        layout.addRow(QLabel("Script:"), self.cb_script)

        self.gnc_file_btn = QPushButton("Get Gnucash file")
        self.gnc_file_btn.clicked.connect(partial(self.open_file_name_dialog))
        layout.addRow(QLabel("Gnucash File:"), self.gnc_file_btn)

        self.cb_mode = QComboBox()
        self.cb_mode.addItems([TEST,SHEET_1,SHEET_2])
        self.cb_mode.currentIndexChanged.connect(partial(self.selection_change, self.cb_mode, MODE))
        layout.addRow(QLabel(MODE+':'), self.cb_mode)

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

        self.exe_btn = QPushButton("Go!")
        self.exe_btn.clicked.connect(partial(self.button_click))
        layout.addRow(QLabel("EXECUTE:"), self.exe_btn)

        self.gb_main.setLayout(layout)

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName( self, "Get Gnucash Files", osp.join(BASE_GNUCASH_FOLDER, "bak-files" + osp.sep),
                                                    "Gnucash Files (*.gnc *.gnucash);;All Files (*)", options = options )
        if file_name:
            self.gnc_file = file_name
            gnc_file_display = file_name.split('/')[-1]
            self.gnc_file_btn.setText(gnc_file_display)

    def get_log_level(self):
        num, ok = QInputDialog.getInt(self, "Logging Level", "Enter a value (0-60)", value=self.log_level, min=0, max=60)
        if ok:
            self.log_level = num
            ui_lgr.debug(F"logging level changed to {num}.")

    @staticmethod
    def selection_change(cb:QComboBox, label:str):
        ui_lgr.debug(F"ComboBox '{label}' selection changed to '{cb.currentText()}'.")

    @staticmethod
    def run_function(thread_fxn, p_params:list):
        fxn_param = repr(thread_fxn)
        ui_lgr.info(F"starting thread: {fxn_param}")
        if callable(thread_fxn):
            response = thread_fxn(p_params)
        else:
            msg = F"requested function '{fxn_param}' NOT callable?!"
            ui_lgr.warning(msg)
            raise Exception(msg)
        ui_lgr.info(F"finished thread: {fxn_param}")
        return response

    def button_click(self):
        """Assemble the necessary parameters and call each selected update choice in a separate thread."""
        ui_lgr.info(F"Clicked '{self.exe_btn.text()}'.")
        exe = self.cb_script.currentText()
        ui_lgr.info(F"Script is '{exe}'.")

        if not self.gnc_file:
            self.response_box.append(">>> MUST select a Gnucash File!")
            return

        cl_params = ['-g' + self.gnc_file, '-m' + self.cb_mode.currentText(),
                     '-t' + self.cb_domain.currentText(), '-l' + str(self.log_level)]

        if self.ch_ggl.isChecked(): cl_params.append("--ggl_save")
        if self.ch_gnc.isChecked(): cl_params.append("--gnc_save")
        if self.ch_rsp.isChecked(): cl_params.append("--resp_save")
        ui_lgr.info( repr(cl_params) )

        main_run = CHOICE_FXNS[exe]
        ui_lgr.info(f"functions to run = {main_run}")
        if callable(main_run):
            ui_lgr.info(f"Calling update {exe}...")
            response = main_run(cl_params)
        elif isinstance(main_run, list):
            ui_lgr.info(f"updates to run = {exe}")
            # use 'with' to ensure threads are cleaned up properly
            with confut.ThreadPoolExecutor(max_workers = len(main_run)) as executor:
                # send each update function to a separate thread
                running_threads = {executor.submit(self.run_function, fxn, cl_params):fxn for fxn in main_run}
                ui_lgr.info(F"running threads = {repr(running_threads)}")
                for completed_thread in confut.as_completed(running_threads):
                    submitted_fxn = repr(running_threads[completed_thread])
                    try:
                        data = completed_thread.result()
                    except Exception as thr_ex:
                        ui_lgr.warning(F"{submitted_fxn} generated exception: {repr(thr_ex)}")
                        raise thr_ex
                    else:
                        ui_lgr.debug(F"Update function '{submitted_fxn}' has finished with {type(data)} data")
            response = lg_ctrl.get_saved_info()
        else:
            msg = F"Problem with functions??!! '{exe}'"
            ui_lgr.error(msg)
            response = msg

        reply = {"response":response}
        self.response_box.append( json.dumps(reply, indent=4) )

# END class UpdateBudgetUI


def ui_main():
    app = QApplication(argv)
    dialog = UpdateBudgetUI()
    dialog.show()
    app.exec_()


if __name__ == "__main__":
    lg_ctrl = MhsLogger(UpdateBudgetUI.__name__, con_level = UI_DEFAULT_LOG_LEVEL, suffix = DEFAULT_LOG_SUFFIX)
    ui_lgr = lg_ctrl.get_logger()
    ui_main()
    exit()
