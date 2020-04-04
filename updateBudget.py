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
__updated__ = '2020-04-03'

from sys import path, exc_info
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import *


def process_args(base_year:int) -> ArgumentParser:
    arg_parser = ArgumentParser(description = 'Update the Revenues & Expenses section of my Google Sheet',
                                prog = 'updateRevExps.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required = True, help = 'path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required = True, choices = [TEST, SEND + '1', SEND + '2'],
                          help = 'SEND to Google sheet (1 or 2) OR just TEST')
    required.add_argument('-y', '--year', required = True, help = F"year to update: {base_year}..2019")
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices = ['1', '2', '3', '4'], help = "quarter to update: 1..4")
    arg_parser.add_argument('-l', '--level', type = int, default = lg.INFO, help = 'set LEVEL of logging output')
    arg_parser.add_argument('--gnc_save', action = 'store_true', help = 'Write the Gnucash data to a JSON file')
    arg_parser.add_argument('--ggl_save', action = 'store_true', help = 'Write the Google data to a JSON file')
    arg_parser.add_argument('--resp_save', action = 'store_true', help = 'Write the Google RESPONSE to a JSON file')

    return arg_parser


class UpdateBudget:
    """update my 'Budget Quarterly' Google spreadsheet with information from a Gnucash file"""
    def __init__(self, args:list, p_log_name:str, p_year:int, p_yr_span:int, p_qtr_span:int, p_lgr:lg.Logger):
        self._lgr = p_lgr
        self.base_log_name = p_log_name

        self.process_input_parameters(args, p_year)
        self.base_year = p_year
        self.year_span = p_yr_span
        self.qtr_span = p_qtr_span

        # get info for log names
        _, fname = osp.split(self.gnucash_file)
        base_name, _ = osp.splitext(fname)
        self.target_name = F"-{self.year}{('-Q' + str(self.qtr) if self.qtr else '')}"
        self.log_name = get_logger_filename(p_log_name) + '_' + base_name + self.target_name

        self.now = dt.now().strftime(FILE_DATE_FORMAT)

        self.gnucash_data = []

        self._lgr.setLevel(self.level)
        self._lgr.log(self.level, F"\n\t\tRuntime = {self.now}")

    def get_mode(self) -> str:
        return self.mode

    # noinspection PyAttributeOutsideInit
    def process_input_parameters(self, argl:list, p_year:int):
        args = process_args(p_year).parse_args(argl)
        self._lgr.info(F"\nargs = {args}")

        self._lgr.info(F"logger level set to {args.level}")

        if not osp.isfile(args.gnucash_file):
            msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
            self._lgr.error(msg)
            raise Exception(msg)
        self.gnucash_file = args.gnucash_file
        self._lgr.info(F"\n\t\tGnucash file = {self.gnucash_file}")

        self.year = get_int_year(args.year, p_year)
        self.qtr = 0 if args.quarter is None else get_int_quarter(args.quarter)

        self.save_gnc = args.gnc_save
        self.save_ggl = args.ggl_save
        self.save_resp = args.resp_save
        self.level = args.level
        self.mode = args.mode

    # noinspection PyAttributeOutsideInit
    def prepare_gnucash_data(self, call_object:object):
        """
        Get REVENUE and EXPENSE data for ONE specified Quarter or ALL four Quarters for the specified Year
        >> NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
           the Gnucash data in a separate dict instead of just preparing a Google data dict
           :param call_object: instance with required functions
        """
        # get either ONE Quarter or ALL Quarters if updating an entire Year
        num_quarters = 1 if self.qtr else 4
        self._lgr.info("find Revenue & Expenses in {} for {}{}"
                       .format(self.gnucash_file, self.year, ('-Q' + str(self.qtr)) if self.qtr else ''))
        try:
            self.gnc_session = GnucashSession(self.mode, self.gnucash_file, BOTH, self._lgr)
            self.gnc_session.begin_session()

            for i in range(num_quarters):
                qtr = self.qtr if self.qtr else i + 1
                start_month = (qtr * 3) - 2

                # for each period keep the start date, end date, debits and credits sums and overall total
                period_list = [
                    [
                        start_date, end_date,
                        ZERO, # debits sum
                        ZERO, # credits sum
                        ZERO  # TOTAL
                    ]
                    for start_date, end_date in generate_quarter_boundaries(self.year, start_month, 1)
                ]
                # a copy of the above list with just the period start dates
                period_starts = [e[0] for e in period_list]

                data_quarter = {}
                call_object.get_data(self.gnc_session.get_root_acct(), qtr, period_starts, period_list, self.year, data_quarter)
                self._lgr.debug(F"\nTOTAL Expenses for {self.year}-Q{qtr} = {period_list[0][4]}\n")

                self.gnucash_data.append(data_quarter)
                self._lgr.log(5, json.dumps(data_quarter, indent=4))

            # no save needed, we're just reading...
            self.gnc_session.end_session(False)

            if self.save_gnc:
                fname = F"updateRevExps_gnc-data-{self.year}{('-Q' + str(self.qtr) if self.qtr else '')}"
                self._lgr.info(F"gnucash data file = {save_to_json(fname, self.gnucash_data)}")

        except Exception as fgde:
            fgde_msg = F"prepare_gnucash_data() EXCEPTION: {repr(fgde)}!"
            tb = exc_info()[2]
            self._lgr.error(fgde_msg, tb)
            if self.gnc_session:
                self.gnc_session.check_end_session(locals())
            raise fgde.with_traceback(tb)

    @staticmethod
    def fill_google_cell(call_object:object, p_dest:str, p_col:str, p_row:int, p_time:str):
        call_object.gglu.fill_cell(p_dest, p_col, p_row, p_time)

    def fill_google_data(self, call_object:object):
        """
        Fill the data list:
        for each item in results, either 1 for one quarter or 4 for four quarters:
        create 5 cell_data's, one each for REV, BAL, CONT, NEC, DEDNS:
        fill in the range based on the year and quarter
        range = SHEET_NAME + '!' + calculated cell
        fill in the values based on the sheet being updated and the type of cell_data
        REV string is '= ${INV} + ${OTH} + ${SAL}'
        DEDNS string is '= ${Mk-Dedns} + ${Lu-Dedns} + ${ML-Dedns}'
        others are just the string from the item
        """
        self._lgr.info(get_current_time())
        year_row = BASE_ROW + year_span(self.year, self.base_year, self.year_span, 0)
        # get exact row from Quarter value in each item
        for item in self.gnucash_data:
            self._lgr.log(5, F"{QTR} = {item[QTR]}")
            dest_row = year_row + ((get_int_quarter(item[QTR]) - 1) * self.qtr_span)
            self._lgr.log(5, F"dest_row = {dest_row}\n")
            for key in item:
                if key != QTR:
                    dest = call_object.nec_inc_dest
                    if key in (REV, BAL, CONT):
                        dest = call_object.all_inc_dest
                    self.fill_google_cell(call_object, dest, REV_EXP_COLS[key], dest_row, item[key])

        self.fill_update_date_and_time(call_object)

        str_qtr = None
        if len(self.gnucash_data) == 1:
            str_qtr = self.gnucash_data[0][QTR]

        if self.save_ggl:
            fname = F"updateRevExps_google-data-{str(self.year)}{('-Q' + str_qtr if str_qtr else '')}"
            self._lgr.info(F"google data file = {save_to_json(fname, call_object.get_google_data())}")

    # TODO: keep record of all changes: what exactly and when
    def fill_update_date_and_time(self, call_object:object):
        """fill update date & time to ALL and NEC"""
        today_row = BASE_ROW - 1 + year_span(now_dt.year + 2, self.base_year, self.year_span, 0)
        self.fill_google_cell(call_object, call_object.nec_inc_dest, REV_EXP_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(call_object, call_object.nec_inc_dest, REV_EXP_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))
        self.fill_google_cell(call_object, call_object.all_inc_dest, REV_EXP_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(call_object, call_object.all_inc_dest, REV_EXP_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))

    def go(self, main_obj:object) -> dict:
        try:
            # READ the required Gnucash data
            self.prepare_gnucash_data(main_obj)

            # package the Gnucash data in the update format required by Google sheets
            self.fill_google_data(main_obj)

            # check if SENDING data
            if SEND in self.mode:
                response = main_obj.gglu.send_sheets_data()
                if self.save_resp:
                    rf_name = F"UpdateRevExps_response{self.target_name}"
                    self._lgr.info(F"google response file = {save_to_json(rf_name, response, self.now)}")
            else:
                response = {'Response':saved_log_info}

        except Exception as reme:
            reme_msg = repr(reme)
            self._lgr.error(reme_msg)
            response = {'update_rev_exps_main() EXCEPTION':F"{reme_msg}"}

        self._lgr.info(" >>> PROGRAM ENDED.\n")
        finish_logging(self.base_log_name, self.log_name, self.now)
        return response

# END class UpdateBudget
