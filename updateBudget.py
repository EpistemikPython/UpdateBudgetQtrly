##############################################################################################################################
# coding=utf-8
#
# updateBudget.py -- common functions used by the UpdateGoogleSheet project
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2020-03-31'
__updated__ = '2020-04-05'

from sys import path, exc_info
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import *

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
    required.add_argument('-m', '--mode', required = True, choices = [TEST, SEND + '1', SEND + '2'],
                          help = 'SEND to Google sheet (1 or 2) OR just TEST')
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
    """update my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file"""
    def __init__(self, args:list, p_log_name:str, p_sheet_data:dict):
        self._lgr = get_logger(p_log_name)
        self._lgr.info(F"{self.__class__.__name__}({p_log_name})")
        self.base_log_name = p_log_name

        self.base_data = p_sheet_data
        self.process_input_parameters(args, p_sheet_data.get(BASE_YEAR))

        # get info for log names
        _, fname = osp.split(self._gnucash_file)
        base_name, _ = osp.splitext(fname)
        self.target_name = F"-{self.domain}"
        self.log_name = get_logger_filename(p_log_name) + '_' + base_name + self.target_name

        self.now = dt.now().strftime(FILE_DATE_FORMAT)

        self._gnucash_data = []

        self._lgr.setLevel(self.level)
        self._lgr.info(F"\n\t\tRuntime = {self.now}")

    def get_logger(self) -> lg.Logger:
        return self._lgr

    def get_mode(self) -> str:
        return self.mode

    def get_gnucash_file(self) -> str:
        return self._gnucash_file

    # noinspection PyAttributeOutsideInit
    def process_input_parameters(self, argl:list, p_year:int):
        args = process_args(p_year).parse_args(argl)
        # self._lgr.info(F"\nargs = {args}")

        self._lgr.info(F"logger level set to {args.level}")

        if not osp.isfile(args.gnucash_file):
            msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
            self._lgr.error(msg)
            raise Exception(msg)
        self._gnucash_file = args.gnucash_file
        self._lgr.info(F"\n\t\tGnucash file = {self._gnucash_file}")

        self.domain = args.timeframe
        try:
            self.year = get_int_year(self.domain, self.base_data.get(BASE_YEAR))
        except Exception as e:
            msg = repr(e)
            self._lgr.warning(msg)
            self.year = self.base_data.get(BASE_YEAR)

        self.save_gnc = args.gnc_save
        self.save_ggl = args.ggl_save
        self.save_resp = args.resp_save
        self.level = args.level
        self.mode = args.mode

        self._lgr.info(F"year = {self.domain} & mode = {self.mode}")

    # noinspection PyAttributeOutsideInit
    def prepare_gnucash_data(self, call_object:object):
        """
        Get data for the specified year, or ALL years
            NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
            the Gnucash data in a separate dict instead of just preparing a Google data dict
        :param call_object: instance with required functions
        """
        # for ALL Quarters since updating an entire Year
        num_quarters = 4
        self._lgr.info(F"call object = {str(call_object)}")
        gnc_session = None
        try:
            gnc_session = GnucashSession(self.mode, self._gnucash_file, BOTH, self._lgr)
            gnc_session.begin_session()

            for i in range(num_quarters):
                data_quarter = {}
                call_object.fill_gnucash_data(gnc_session, i+1, self.year, data_quarter)

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

    def fill_google_data(self, call_object:object):
        """ Fill the Google data list """
        self._lgr.info(get_current_time())

        call_object.fill_google_data(self.domain)

        self.record_update(call_object)

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

    def go(self, update_subtype:object) -> dict:
        try:
            # READ the required Gnucash data
            self.prepare_gnucash_data(update_subtype)

            # package the Gnucash data in the update format required by Google sheets
            self.fill_google_data(update_subtype)

            # check if SENDING data
            if SEND in self.mode:
                response = update_subtype.send_sheets_data()
                if self.save_resp:
                    rf_name = F"{update_subtype.__class__.__name__}_response{self.target_name}"
                    self._lgr.info(F"google response file = {save_to_json(rf_name, response, self.now)}")
            else:
                response = {'Response':saved_log_info}

        except Exception as goe:
            goe_msg = repr(goe)
            self._lgr.error(goe_msg)
            response = {'go() EXCEPTION':F"{goe_msg}"}

        self._lgr.info(" >>> PROGRAM ENDED.\n")
        finish_logging(self.base_log_name, self.log_name, self.now)
        return response

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
