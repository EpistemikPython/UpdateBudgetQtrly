##############################################################################################################################
# coding=utf-8
#
# updateBudget.py -- common functions used by the UpdateGoogleSheet project
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>

__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2020-03-31'
__updated__ = '2020-06-07'

from sys import path, exc_info
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import *

UPDATE_YEARS   = ['2020', '2019', '2018', '2017', '2016', '2015', '2014', '2013', '2012', '2011', '2010', '2009', '2008']
UPDATE_DOMAINS = copy(UPDATE_YEARS)
UPDATE_DOMAINS.append(ALL_YRS)
print(F"Update Domains = {UPDATE_DOMAINS}")

SHEET_1:str   = SHEET + ' 1'
SHEET_2:str   = SHEET + ' 2'
BASE_YEAR:str = BASE + YR
YEAR_SPAN:str = BASE_YEAR + SPAN
QTR_SPAN:str  = QTR + SPAN
HDR_SPAN:str  = 'Header' + SPAN

RECORD_RANGE    = "'Record'!A1"
RECORD_SHEET    = 'Record'
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
    required.add_argument('-t', '--timeframe', required=True,
                          help=F"'today' | 'current year' | 'previous year' | {base_year}..{now_dt.year} | 'allyears'")
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices = ['1', '2', '3', '4'], help = "quarter to update: 1..4")
    arg_parser.add_argument('-l', '--level', type = int, default = lg.INFO, help = 'set LEVEL of logging output')
    arg_parser.add_argument('--gnc_save', action = 'store_true', help = 'Write the Gnucash data to a JSON file')
    arg_parser.add_argument('--ggl_save', action = 'store_true', help = 'Write the Google data to a JSON file')
    arg_parser.add_argument('--resp_save', action = 'store_true', help = 'Write the Google RESPONSE to a JSON file')

    return arg_parser


class UpdateBudget:
    """
    update my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file
    -- contains common code for the three options of updating rev&exps, assets, balances
    """
    def __init__(self, args:list, p_log_name:str, p_sheet_data:dict):
        self._lgr = get_logger(p_log_name)
        self._lgr.info(F"init: {self.__class__.__name__}({p_log_name})")
        self.base_log_name = p_log_name

        self.base_data = p_sheet_data
        self.process_input_parameters(args, p_sheet_data.get(BASE_YEAR))

        # get info for log names
        _, fname = osp.split(self._gnucash_file)
        base_name, _ = osp.splitext(fname)
        self.target_name = F"-{self.domain}"
        self.log_name = get_logger_filename(p_log_name) + '_' + base_name + self.target_name

        self._lgr.setLevel(self.level)
        self._lgr.info(F"\n\t\t{self.__class__.__name__} for {self._gnucash_file}:\n\t\tRuntime = {get_current_time()}")

        self._gnucash_data = []

        self._gt = None
        self.response = None

    def get_logger(self) -> lg.Logger:
        return self._lgr

    def get_mode(self) -> str:
        return self.mode

    def get_gnucash_file(self) -> str:
        return self._gnucash_file

    # noinspection PyAttributeOutsideInit
    def process_input_parameters(self, argl:list, p_year:int):
        args = process_args(p_year).parse_args(argl)

        if not osp.isfile(args.gnucash_file):
            msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
            self._lgr.error(msg)
            raise Exception(msg)
        self._gnucash_file = args.gnucash_file

        self.domain = args.timeframe
        self.level  = args.level
        self.mode   = args.mode

        self.save_gnc  = args.gnc_save
        self.save_ggl  = args.ggl_save
        self.save_resp = args.resp_save

        self._lgr.info(F"\n\t\tGnucash file = {self._gnucash_file}\n\t\tDomain = {self.domain} & Mode = {self.mode}")

    # noinspection PyAttributeOutsideInit
    def prepare_gnucash_data(self, call_object:object, p_years:list):
        """
        Get data for the specified year, or ALL years
            NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
            the Gnucash data in a separate dict instead of just preparing a Google data dict
        :param call_object: instance with required functions
        :param p_years: year(s) to update
        """
        self._lgr.info(F"call object '{str(call_object)}' at {get_current_time()}")
        gnc_session = None
        try:
            gnc_session = GnucashSession(self.mode, self._gnucash_file, BOTH, self._lgr)
            gnc_session.begin_session()

            for year in p_years:
                for i in range(4): # ALL quarters since updating an entire year
                    data_quarter = {}
                    call_object.fill_gnucash_data(gnc_session, i+1, year, data_quarter)

                    self._gnucash_data.append(data_quarter)
                    self._lgr.debug(json.dumps(data_quarter, indent=4))

            # no save needed, we're just reading...
            gnc_session.end_session(False)

            if self.save_gnc:
                fname = F"{call_object.__class__.__name__}_gnc-data-{self.domain}"
                self._lgr.info(F"gnucash data file = {save_to_json(fname, self._gnucash_data)}")

        except Exception as fgde:
            fgde_msg = F"prepare_gnucash_data() EXCEPTION: {repr(fgde)}!"
            tb = exc_info()[2]
            self._lgr.error(fgde_msg, tb)
            if gnc_session:
                gnc_session.check_end_session(locals())
            raise fgde.with_traceback(tb)

    def prepare_google_data(self, call_object:object, p_years:list):
        """fill the Google data list"""
        self._lgr.info(F"call object '{str(call_object)}' at {get_current_time()}")

        for year in p_years:
            call_object.fill_google_data(year)

        if self.save_ggl:
            fname = F"{call_object.__class__.__name__}_google-data-{str(self.domain)}"
            self._lgr.info(F"google data file = {save_to_json(fname, call_object.get_google_data())}")

    def record_update(self, call_object:object):
        ggl_updater = call_object.get_google_updater()
        ru_result = ggl_updater.read_sheets_data(RECORD_RANGE)
        current_row = int(ru_result[0][0])
        self._lgr.info(F"current row = {current_row}\n")

        update_info = call_object.__class__.__name__ + ' - ' + self.domain + ' - ' + self.get_mode()
        self._lgr.info(F"update info = {update_info}\n")

        # keep record of this update
        ggl_updater.fill_cell(RECORD_SHEET, RECORD_DATE_COL, current_row, now_dt.strftime(CELL_DATE_STR))
        ggl_updater.fill_cell(RECORD_SHEET, RECORD_TIME_COL, current_row, now_dt.strftime(CELL_TIME_STR))
        ggl_updater.fill_cell(RECORD_SHEET, RECORD_GNC_COL,  current_row, self.get_gnucash_file())
        ggl_updater.fill_cell(RECORD_SHEET, RECORD_INFO_COL, current_row, update_info)

        # update the row tally
        ggl_updater.fill_cell(RECORD_SHEET, RECORD_DATE_COL, 1, str(current_row+1))

    def start_google_thread(self, call_object:object):
        self._lgr.info("before creating thread")
        self._gt = threading.Thread(target = self.send_google_data, args = (call_object,))
        self._lgr.info("before running thread")
        self._gt.start()
        self._lgr.info(F"thread '{str(call_object)}' started at {get_current_time()}")

    def send_google_data(self, call_object:object):
        call_object.get_google_updater().begin_session()

        self.record_update(call_object)
        self.response = call_object.send_sheets_data()

        call_object.get_google_updater().end_session()

        if self.save_resp:
            rf_name = F"{call_object.__class__.__name__}_response{self.target_name}"
            self._lgr.info(F"google response file = "
                           F"{save_to_json(rf_name, self.response, get_current_time(FILE_DATETIME_FORMAT))}")

    def go(self, update_subtype:object) -> dict:
        """
        starting point for accessing UpdateBudget functions
        """
        years = UPDATE_YEARS if self.domain == ALL_YRS else [self.domain]
        try:
            # READ the required Gnucash data
            self.prepare_gnucash_data(update_subtype, years)

            # package the Gnucash data in the update format required by Google sheets
            self.prepare_google_data(update_subtype, years)

            # check if SENDING data
            if SHEET in self.mode:
                self.start_google_thread(update_subtype)
            else:
                self.response = {'Response':saved_log_info}

        except Exception as goe:
            goe_msg = repr(goe)
            self._lgr.error(goe_msg)
            self.response = {'go() EXCEPTION':F"{goe_msg}"}

        # check if we started the google thread and wait if necessary
        if self._gt and self._gt.is_alive():
            self._lgr.info("wait for the thread to finish")
            self._gt.join()
        self._lgr.info(" >>> PROGRAM ENDED.\n")
        finish_logging(self.base_log_name, self.log_name, get_current_time(FILE_DATETIME_FORMAT), sfx='gncout')
        return self.response

# END class UpdateBudget


def test_google_read():
    logger = get_logger(UpdateBudget.__name__)
    ggl_updater = GoogleUpdate(logger)
    result = ggl_updater.read_sheets_data(RECORD_RANGE)
    print(result)
    print(result[0][0])


if __name__ == "__main__":
    test_google_read()
    exit()
