##############################################################################################################################
# coding=utf-8
#
# updateBudget.py -- common functions used by the UpdateGoogleSheet project
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>

__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2020-03-31'
__updated__ = '2020-08-03'

from sys import path, exc_info
from abc import ABC, abstractmethod
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import *

UPDATE_YEARS = ['2020', '2019', '2018', '2017', '2016', '2015', '2014', '2013', '2012', '2011', '2010', '2009', '2008']
BASE_UPDATE_YEAR = UPDATE_YEARS[-1]

SHEET_1:str   = SHEET + " 1"
SHEET_2:str   = SHEET + " 2"
BASE_YEAR:str = BASE + YR
YEAR_SPAN:str = BASE_YEAR + SPAN
QTR_SPAN:str  = QTR + SPAN
HDR_SPAN:str  = 'Header' + SPAN

RECORD_RANGE    = "'Record'!A1"
RECORD_SHEET    = "Record"
RECORD_DATE_COL = 'A'
RECORD_TIME_COL = 'B'
RECORD_GNC_COL  = 'C'
RECORD_INFO_COL = 'D'


def process_args(base_year:int) -> ArgumentParser:
    arg_parser = ArgumentParser(description = 'Update the Revenues & Expenses section of my Google Sheet',
                                prog = 'updateRevExps.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required = True, help = 'path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required = True, choices = [TEST, SHEET_1, SHEET_2],
                          help = 'SEND to Google Sheet (1 or 2) OR just TEST')
    required.add_argument('-t', '--timespan', required=True,
        help=F"choices = {CURRENT_YRS} | {RECENT_YRS} | {MID_YRS} | {EARLY_YRS}| {ALL_YRS} | {base_year}..{now_dt.year}")
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices = ['1', '2', '3', '4'], help = "quarter to update: 1..4")
    arg_parser.add_argument('-l', '--level', type = int, default = lg.INFO, help = 'set LEVEL of logging output')
    arg_parser.add_argument('--gnc_save', action = 'store_true', help = 'Write the Gnucash data to a JSON file')
    arg_parser.add_argument('--ggl_save', action = 'store_true', help = 'Write the Google data to a JSON file')
    arg_parser.add_argument('--resp_save', action = 'store_true', help = 'Write the Google RESPONSE to a JSON file')

    return arg_parser


def get_timespan(timespan:str, lgr:lg.Logger) -> list:
    if timespan == ALL_YRS:
        return UPDATE_YEARS
    elif timespan == EARLY_YRS:
        return UPDATE_YEARS[10:13]
    elif timespan == MID_YRS:
        return UPDATE_YEARS[6:10]
    elif timespan == RECENT_YRS:
        return UPDATE_YEARS[2:6]
    elif timespan == CURRENT_YRS:
        return UPDATE_YEARS[0:2]
    elif timespan in UPDATE_YEARS:
        return [timespan]
    lgr.warning(F"INVALID YEAR: {timespan}")
    return UPDATE_YEARS[0:1]


class UpdateBudget(ABC):
    """
    update my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file
    -- contains common code for the three options of updating rev&exps, assets, balances
    """
    def __init__(self, args:list, p_log_name:str, p_base_year:int):
        self._lgr = get_logger(p_log_name)

        self.process_input_parameters(args, p_base_year)

        # get info for log names
        self.base_log_name = p_log_name
        _, fname = osp.split(self._gnucash_file)
        base_name, _ = osp.splitext(fname)
        self.target_name = F"-{self.timespan}"
        self.log_name = get_logger_filename(p_log_name) + '_' + base_name + self.target_name

        self._lgr.setLevel(self.level)
        self._lgr.info(F"\n\t\t{self.__class__.__name__}({p_log_name}) for {self._gnucash_file}:"
                       F"\n\t\tRuntime = {get_current_time()}")

        self._gnucash_data = []
        self._ggl_update   = GoogleUpdate(self._lgr)

        self._ggl_thrd = None
        self.response  = None

    # noinspection PyAttributeOutsideInit
    def process_input_parameters(self, argl:list, p_year:int):
        args = process_args(p_year).parse_args(argl)

        if not osp.isfile(args.gnucash_file):
            msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
            self._lgr.error(msg)
            raise Exception(msg)
        self._gnucash_file = args.gnucash_file

        self.timespan = args.timespan
        self.level    = args.level
        self.mode     = args.mode

        self.save_gnc  = args.gnc_save
        self.save_ggl  = args.ggl_save
        self.save_resp = args.resp_save

        self._lgr.info(F"\n\t\tGnucash file = {self._gnucash_file}\n\t\tDomain = {self.timespan} & Mode = {self.mode}")

    # noinspection PyAttributeOutsideInit
    def prepare_gnucash_data(self, p_years:list):
        """
        Get data for the specified year, or group of years
            NOT really necessary to create a collection of the Gnucash data, but useful to store all
            the Gnucash data in a separate dict instead of just directly preparing a Google data dict
        :param p_years: year(s) to update
        """
        self._lgr.info(F"prepare_gnucash_data({p_years}) at {get_current_time()}")
        gnc_session = None
        try:
            gnc_session = GnucashSession(self.mode, self._gnucash_file, BOTH, self._lgr)
            gnc_session.begin_session()

            for year in p_years:
                for i in range(4): # ALL quarters since updating an entire year
                    data_quarter = {}
                    self.fill_gnucash_data(gnc_session, i+1, year, data_quarter)

                    self._gnucash_data.append(data_quarter)
                    self._lgr.debug(json.dumps(data_quarter, indent=4))

            # no save needed, we're just reading...
            gnc_session.end_session(False)

            if self.save_gnc:
                fname = F"{self.__class__.__name__}_gnc-data-{self.timespan}"
                self._lgr.info(F"gnucash data file = {save_to_json(fname, self._gnucash_data)}")

        except Exception as fgde:
            fgde_msg = F"prepare_gnucash_data() EXCEPTION: {repr(fgde)}!"
            tb = exc_info()[2]
            self._lgr.error(fgde_msg, tb)
            if gnc_session:
                gnc_session.check_end_session(locals())
            raise fgde.with_traceback(tb)

    def prepare_google_data(self, p_years:list):
        """fill the Google data list"""
        self._lgr.info(F"prepare_google_data({p_years}) at {get_current_time()}")

        self.fill_google_data(p_years)

        if self.save_ggl:
            fname = F"{self.__class__.__name__}_google-data-{str(self.timespan)}"
            self._lgr.info(F"google data file = {save_to_json(fname, self._ggl_update.get_data())}")

    def record_update(self):
        ru_result = self._ggl_update.read_sheets_data(RECORD_RANGE)
        current_row = int(ru_result[0][0])
        self._lgr.info(F"current row = {current_row}\n")

        update_info = self.__class__.__name__ + ' - ' + self.timespan + ' - ' + self.mode
        self._lgr.info(F"update info = {update_info}\n")

        # keep record of this update
        self._ggl_update.fill_cell(RECORD_SHEET, RECORD_DATE_COL, current_row, now_dt.strftime(CELL_DATE_STR))
        self._ggl_update.fill_cell(RECORD_SHEET, RECORD_TIME_COL, current_row, now_dt.strftime(CELL_TIME_STR))
        self._ggl_update.fill_cell(RECORD_SHEET, RECORD_GNC_COL, current_row, self._gnucash_file)
        self._ggl_update.fill_cell(RECORD_SHEET, RECORD_INFO_COL, current_row, update_info)

        # update the row tally
        self._ggl_update.fill_cell(RECORD_SHEET, RECORD_DATE_COL, 1, str(current_row + 1))

    def start_google_thread(self):
        self._lgr.info("before creating thread")
        self._ggl_thrd = threading.Thread(target = self.send_google_data)
        self._lgr.info("before running thread")
        self._ggl_thrd.start()
        self._lgr.info(F"thread '{str(self)}' started at {get_current_time()}")

    def send_google_data(self):
        self._ggl_update.begin_session()

        self.record_update()
        self.response = self._ggl_update.send_sheets_data()

        self._ggl_update.end_session()

        if self.save_resp:
            rf_name = F"{self.__class__.__name__}_response{self.target_name}"
            self._lgr.info(F"google response file = "
                           F"{save_to_json(rf_name, self.response, get_current_time(FILE_DATETIME_FORMAT))}")

    def go(self) -> dict:
        """
        ENTRY POINT for accessing UpdateBudget functions
        """
        years = get_timespan(self.timespan, self._lgr)
        self._lgr.info(F"timespan to find = {years}")
        try:
            # READ the required Gnucash data
            self.prepare_gnucash_data(years)

            # package the Gnucash data in the update format required by Google sheets
            self.prepare_google_data(years)

            # check if SENDING data
            if SHEET in self.mode:
                self.start_google_thread()
            else:
                self.response = {'Response' : saved_log_info}

        except Exception as goe:
            goe_msg = repr(goe)
            self._lgr.error(goe_msg)
            self.response = {'go() EXCEPTION' : F"{goe_msg}"}

        # check if we started the google thread and wait if necessary
        if self._ggl_thrd and self._ggl_thrd.is_alive():
            self._lgr.info("wait for the thread to finish")
            self._ggl_thrd.join()
        self._lgr.info(" >>> PROGRAM ENDED.\n")
        finish_logging(self.base_log_name, self.log_name, get_current_time(FILE_DATETIME_FORMAT), sfx='gncout')
        return self.response

    @abstractmethod
    def fill_gnucash_data(self, gnc_session, param, year, data_quarter):
        pass

    @abstractmethod
    def fill_google_data(self, p_years):
        pass

# END class UpdateBudget


def test_google_read():
    logger = get_logger(UpdateBudget.__name__)
    ggl_updater = GoogleUpdate(logger)
    result = ggl_updater.test_read(RECORD_RANGE)
    print(result)
    print(result[0][0])


if __name__ == "__main__":
    process_args(BASE_UPDATE_YEAR)
    test_google_read()
    exit()
