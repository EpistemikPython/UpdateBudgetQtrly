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
__updated__ = '2020-04-04'

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
    def __init__(self, args:list, p_log_name:str, p_sheet_data:dict, p_lgr:lg.Logger):
        self._lgr = p_lgr
        self.base_log_name = p_log_name

        self.base_data = p_sheet_data
        self.process_input_parameters(args, p_sheet_data.get(BASE_YEAR))

        # get info for log names
        _, fname = osp.split(self.gnucash_file)
        base_name, _ = osp.splitext(fname)
        self.target_name = F"-{self.year}{('-Q' + str(self.qtr) if self.qtr else '')}"
        self.log_name = get_logger_filename(p_log_name) + '_' + base_name + self.target_name

        self.now = dt.now().strftime(FILE_DATE_FORMAT)

        self.gnucash_data = []

        self._lgr.setLevel(self.level)
        self.info(F"\n\t\tRuntime = {self.now}")

    def get_mode(self) -> str:
        return self.mode

    # noinspection PyAttributeOutsideInit
    def process_input_parameters(self, argl:list, p_year:int):
        args = process_args(p_year).parse_args(argl)
        self.info(F"\nargs = {args}")

        self.info(F"logger level set to {args.level}")

        if not osp.isfile(args.gnucash_file):
            msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
            self.error(msg)
            raise Exception(msg)
        self.gnucash_file = args.gnucash_file
        self.info(F"\n\t\tGnucash file = {self.gnucash_file}")

        self.year = get_int_year(args.year, p_year)
        self.qtr  = 0 if args.quarter is None else get_int_quarter(args.quarter)
        self.info(F"year = {self.year} & quarter = {self.qtr}")

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
        self.info("find Revenue & Expenses in {} for {}{}"
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
                call_object.fill_gnucash_data(self.gnc_session.get_root_acct(), qtr, period_starts, period_list,
                                              self.year, data_quarter)
                self.debug(F"\nTOTAL Expenses for {self.year}-Q{qtr} = {period_list[0][4]}\n")

                self.gnucash_data.append(data_quarter)
                self.debug(json.dumps(data_quarter, indent=4))

            # no save needed, we're just reading...
            self.gnc_session.end_session(False)

            if self.save_gnc:
                fname = F"updateRevExps_gnc-data-{self.year}{('-Q' + str(self.qtr) if self.qtr else '')}"
                self.info(F"gnucash data file = {save_to_json(fname, self.gnucash_data)}")

        except Exception as fgde:
            fgde_msg = F"prepare_gnucash_data() EXCEPTION: {repr(fgde)}!"
            tb = exc_info()[2]
            self.error(fgde_msg, tb)
            if self.gnc_session:
                self.gnc_session.check_end_session(locals())
            raise fgde.with_traceback(tb)

    def fill_google_data(self, call_object:object):
        """ Fill the Google data list """
        self.info(get_current_time())
        year_row = BASE_ROW + year_span(self.year, self.base_data.get(BASE_YEAR), self.base_data.get(YEAR_SPAN),
                                        self.base_data.get(HDR_SPAN), self._lgr)

        call_object.fill_google_data(year_row)

        self.record_update(call_object)

        if self.save_ggl:
            str_qtr = self.gnucash_data[0][QTR] if len(self.gnucash_data) == 1 else None
            fname = F"updateRevExps_google-data-{str(self.year)}{('-Q' + str_qtr if str_qtr else '')}"
            self.info(F"google data file = {save_to_json(fname, call_object.get_google_data())}")

    def info(self, msg:str): self._lgr.info(msg)
    def debug(self, msg:str): self._lgr.debug(msg)
    def error(self, msg:str, tb:object=None): self._lgr.error(msg, tb)

    # TODO: keep record of all changes to Google sheet: what exactly and when
    @staticmethod
    def record_update(call_object:object):
        call_object.record_update()

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
                    rf_name = F"UpdateRevExps_response{self.target_name}"
                    self.info(F"google response file = {save_to_json(rf_name, response, self.now)}")
            else:
                response = {'Response':saved_log_info}

        except Exception as goe:
            goe_msg = repr(goe)
            self.error(goe_msg)
            response = {'go() EXCEPTION':F"{goe_msg}"}

        self.info(" >>> PROGRAM ENDED.\n")
        finish_logging(self.base_log_name, self.log_name, self.now)
        return response

# END class UpdateBudget
