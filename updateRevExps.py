##############################################################################################################################
# coding=utf-8
#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Revenue and Expenses
#                     in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__  = 3.9
__gnucash_version__ = 3.8
__created__ = '2019-03-30'
__updated__ = '2020-01-11'

from sys import path, exc_info
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import GoogleUpdate, BASE_ROW

# path to the account in the Gnucash file
REV_ACCTS = {
    INV : ["REV_Invest"],
    OTH : ["REV_Other"],
    EMPL: ["REV_Employment"]
}
EXP_ACCTS = {
    BAL  : ["EXP_Balance"],
    CONT : ["EXP_CONTINGENT"],
    NEC  : ["EXP_NECESSARY"]
}
DEDNS_BASE = 'DEDNS_Employment'
DEDN_ACCTS = {
    "Mark" : [DEDNS_BASE, 'Mark'],
    "Lulu" : [DEDNS_BASE, 'Lulu'],
    "ML"   : [DEDNS_BASE, 'Marie-Laure']
}

# column index in the Google sheets
REV_EXP_COLS = {
    DATE  : 'B',
    REV   : 'D',
    BAL   : 'P',
    CONT  : 'O',
    NEC   : 'G',
    DEDNS : 'D'
}

# sheet names in Budget Quarterly
ALL_INC_SHEET:str   = 'All Inc 1'
ALL_INC_2_SHEET:str = 'All Inc 2'
NEC_INC_SHEET:str   = 'Nec Inc 1'
NEC_INC_2_SHEET:str = 'Nec Inc 2'

BOOL_NEC_INC = False
BOOL_ALL_INC = True


class UpdateRevExps:
    def __init__(self, p_filename:str, p_mode:str, p_debug:bool):
        self.debug = p_debug
        self._logger = SattoLog(my_color=BROWN, do_printing=p_debug)
        self._log('UpdateRevExps', GREEN)

        self.gnucash_file = p_filename
        self.gnucash_data = []

        self.gnc_session = None
        self.gglu = GoogleUpdate(self._logger)

        self.mode = p_mode
        # Google sheet to update
        self.all_inc_dest = ALL_INC_2_SHEET
        self.nec_inc_dest = NEC_INC_2_SHEET
        if '1' in self.mode:
            self.all_inc_dest = ALL_INC_SHEET
            self.nec_inc_dest = NEC_INC_SHEET
        self._log(F"all_inc_dest = {self.all_inc_dest}")
        self._log(F"nec_inc_dest = {self.nec_inc_dest}\n")

    BASE_YEAR:int = 2012
    # number of rows between same quarter in adjacent years
    BASE_YEAR_SPAN:int = 11
    # number of rows between quarters in the same year
    QTR_SPAN:int = 2

    def _log(self, p_msg:str, p_color:str=''):
        self._logger.print_info(p_msg, p_color, p_info=inspect.currentframe().f_back)

    def _err(self, p_msg:str, err_info:object):
        self._logger.print_info(p_msg, BR_RED, p_info=err_info)

    def get_gnucash_data(self) -> list:
        return self.gnucash_data

    def get_google_data(self) -> list:
        return self.gglu.get_data()

    def get_log(self) -> list :
        return self._logger.get_log()

    def fill_splits(self, account_path:list, period_starts:list, periods:list):
        return fill_splits(self.gnc_session.get_root_acct(), account_path, period_starts, periods)

    def get_revenue(self, period_starts:list, periods:list) -> dict:
        """
        Get REVENUE data for the specified periods
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :return: revenue for period
        """
        self._log('UpdateRevExps.get_revenue()')
        data_quarter = {}
        str_rev = '= '
        for item in REV_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_base = REV_ACCTS[item]
            acct_name = self.fill_splits(acct_base, period_starts, periods)

            sum_revenue = (periods[0][2] + periods[0][3]) * (-1)
            str_rev += sum_revenue.to_eng_string() + (' + ' if item != EMPL else '')
            self._log(F"{acct_name} Revenue for period = ${sum_revenue}")

        data_quarter[REV] = str_rev
        return data_quarter

    def get_deductions(self, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> str:
        """
        Get SALARY DEDUCTIONS data for the specified Quarter
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param      data_qtr: collected data for the quarter
        :return: deductions for period
        """
        self._log('UpdateRevExps.get_deductions()')
        str_dedns = '= '
        for item in DEDN_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_path = DEDN_ACCTS[item]
            acct_name = self.fill_splits(acct_path, period_starts, periods)

            sum_deductions = periods[0][2] + periods[0][3]
            str_dedns += sum_deductions.to_eng_string() + (' + ' if item != "ML" else '')
            self._log(F"{acct_name} {EMPL} Deductions for {p_year}-Q{data_qtr[QTR]} = ${sum_deductions}")

        data_qtr[DEDNS] = str_dedns
        return str_dedns

    def get_expenses(self, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> str:
        """
        Get EXPENSE data for the specified Quarter
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param      data_qtr: collected data for the quarter
        :return: total expenses for period
        """
        self._log('UpdateRevExps.get_expenses()')
        str_total = ''
        for item in EXP_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_base = EXP_ACCTS[item]
            acct_name = self.fill_splits(acct_base, period_starts, periods)

            sum_expenses = periods[0][2] + periods[0][3]
            str_expenses = sum_expenses.to_eng_string()
            data_qtr[item] = str_expenses
            self._log(F"{acct_name.split('_')[-1]} Expenses for {p_year}-Q{data_qtr[QTR]} = ${str_expenses}")
            str_total += str_expenses + ' + '

        return str_total

    def fill_gnucash_data(self, save_gnc:bool, p_year:int, p_qtr:int):
        """
        Get REVENUE and EXPENSE data for ONE specified Quarter or ALL four Quarters for the specified Year
        >> NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
           the Gnucash data in a separate dict instead of just preparing a Google data dict
        :param save_gnc: true if want to save the Gnucash data to a JSON file
        :param   p_year: year to update
        :param    p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
        :return: nil
        """
        num_quarters = 1 if p_qtr else 4
        self._log("UpdateRevExps.fill_gnucash_data(): find Revenue & Expenses in {} for {}{}"
                  .format(self.gnucash_file, p_year, ('-Q' + str(p_qtr)) if p_qtr else ''))
        try:
            self.gnc_session = GnucashSession(self.mode, self.gnucash_file, self.debug, BOTH)
            self.gnc_session.begin_session()

            for i in range(num_quarters):
                qtr = p_qtr if p_qtr else i + 1
                start_month = (qtr * 3) - 2

                # for each period keep the start date, end date, debits and credits sums and overall total
                period_list = [
                    [
                        start_date, end_date,
                        ZERO, # debits sum
                        ZERO, # credits sum
                        ZERO  # TOTAL
                    ]
                    for start_date, end_date in generate_quarter_boundaries(p_year, start_month, 1)
                ]
                # a copy of the above list with just the period start dates
                period_starts = [e[0] for e in period_list]

                data_quarter = self.get_revenue(period_starts, period_list)
                data_quarter[QTR] = str(qtr)
                self._log(F"\nTOTAL Revenue for {p_year}-Q{qtr} = ${period_list[0][4] * -1}")

                period_list[0][4] = ZERO
                self.get_expenses(period_starts, period_list, p_year, data_quarter)
                self.get_deductions(period_starts, period_list, p_year, data_quarter)
                self._log(F"\nTOTAL Expenses for {p_year}-Q{qtr} = {period_list[0][4]}\n")

                self.gnucash_data.append(data_quarter)
                self._log(json.dumps(data_quarter, indent=4))

            # no save needed, we're just reading...
            self.gnc_session.end_session(False)

            if save_gnc:
                fname = "out/updateRevExps_gnc-data-{}{}".format(p_year, ('-Q' + str(p_qtr)) if p_qtr else '')
                save_to_json(fname, now, self.gnucash_data)

        except Exception as fgde:
            fgde_msg = F"fill_gnucash_data() EXCEPTION: {repr(fgde)}!"
            tb = exc_info()[2]
            self._err(fgde_msg, tb)
            if self.gnc_session:
                self.gnc_session.check_end_session(locals())
            raise fgde.with_traceback(tb)

    def fill_google_cell(self, p_all:bool, p_col:str, p_row:int, p_time:str):
        dest = self.nec_inc_dest
        if p_all:
            dest = self.all_inc_dest
        self.gglu.fill_cell(dest, p_col, p_row, p_time)

    def fill_google_data(self, p_year:int, save_google:bool):
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
        :param      p_year: year to update
        :param save_google: save the Google data to a JSON file
        :return: nil
        """
        self._log('UpdateRevExps.fill_google_data()')
        year_row = BASE_ROW + year_span(p_year, self.BASE_YEAR, self.BASE_YEAR_SPAN, 0)
        # get exact row from Quarter value in each item
        for item in self.gnucash_data:
            self._log(F"{QTR} = {item[QTR]}")
            dest_row = year_row + ((get_int_quarter(item[QTR]) - 1) * self.QTR_SPAN)
            self._log(F"dest_row = {dest_row}\n")
            for key in item:
                if key != QTR:
                    dest = BOOL_NEC_INC
                    if key in (REV, BAL, CONT):
                        dest = BOOL_ALL_INC
                    self.fill_google_cell(dest, REV_EXP_COLS[key], dest_row, item[key])

        # fill update date & time to ALL and NEC
        today_row = BASE_ROW - 1 + year_span(today.year+2, self.BASE_YEAR, self.BASE_YEAR_SPAN, 0)
        self.fill_google_cell(BOOL_NEC_INC, REV_EXP_COLS[DATE], today_row, today.strftime(FILE_DATE_STR))
        self.fill_google_cell(BOOL_NEC_INC, REV_EXP_COLS[DATE], today_row+1, today.strftime(CELL_TIME_STR))
        self.fill_google_cell(BOOL_ALL_INC, REV_EXP_COLS[DATE], today_row, today.strftime(FILE_DATE_STR))
        self.fill_google_cell(BOOL_ALL_INC, REV_EXP_COLS[DATE], today_row+1, today.strftime(CELL_TIME_STR))

        str_qtr = None
        if len(self.gnucash_data) == 1:
            str_qtr = self.gnucash_data[0][QTR]

        if save_google:
            fname = "out/updateRevExps_google-data-{}{}".format(str(p_year), ('-Q' + str_qtr) if str_qtr else '')
            save_to_json(fname, now, self.get_google_data())

# END class UpdateRevExps


def process_args() -> ArgumentParser:
    arg_parser = ArgumentParser(description='Update the Revenues & Expenses section of my Google Sheet', prog='updateRevExps.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST,SEND+'1',SEND+'2'],
                          help='SEND to Google sheet (1 or 2) OR just TEST')
    required.add_argument('-y', '--year', required=True, help="year to update: {}..2019".format(UpdateRevExps.BASE_YEAR))
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices=['1','2','3','4'], help="quarter to update: 1..4")
    arg_parser.add_argument('--gnc_save',  action='store_true', help='Write the Gnucash formatted data to a JSON file')
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google formatted data to a JSON file')
    arg_parser.add_argument('--debug', action='store_true', help='GENERATE DEBUG OUTPUT: MANY LINES!')

    return arg_parser


def process_input_parameters(argl:list) -> (str, bool, bool, bool, str, int, int):
    args = process_args().parse_args(argl)
    SattoLog.print_text(F"\nargs = {args}", BROWN)

    if args.debug:
        SattoLog.print_warning('Printing ALL Debug output!!')

    if not osp.isfile(args.gnucash_file):
        msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
        SattoLog.print_warning(msg)
        raise Exception(msg)
    SattoLog.print_text(F"\nGnucash file = {args.gnucash_file}", GREEN)

    year = get_int_year(args.year, UpdateRevExps.BASE_YEAR)
    qtr = 0 if args.quarter is None else get_int_quarter(args.quarter)

    return args.gnucash_file, args.gnc_save, args.ggl_save, args.debug, args.mode, year, qtr


def update_rev_exps_main(args:list) -> dict :
    SattoLog.print_text(F"Parameters = \n{json.dumps(args, indent=4)}", BROWN)
    gnucash_file, save_gnc, save_ggl, debug, mode, target_year, target_qtr = process_input_parameters(args)

    revexp_now = dt.now().strftime(DATE_STR_FORMAT)
    SattoLog.print_text(F"update_rev_exps_main(): Runtime = {revexp_now}", BLUE)

    try:
        updater = UpdateRevExps(gnucash_file, mode, debug)

        # either for One Quarter or for Four Quarters if updating an entire Year
        updater.fill_gnucash_data(save_gnc, target_year, target_qtr)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_data(target_year, save_ggl)

        # send data if in PROD mode
        if SEND in mode:
            response = updater.gglu.send_sheets_data()
            fname = "out/updateRevExps_response-{}{}".format(target_year , ('-Q' + str(target_qtr)) if target_qtr else '')
            save_to_json(fname, revexp_now, response)
        else:
            response = updater.get_log()

    except Exception as reme:
        reme_msg = repr(reme)
        tb = exc_info()[2]
        SattoLog.print_warning(reme_msg, tb)
        response = {"update_rev_exps_main() EXCEPTION": F"{reme_msg}"}

    SattoLog.print_text(' >>> PROGRAM ENDED.\n', GREEN)
    return response


if __name__ == "__main__":
    from sys import argv
    update_rev_exps_main(argv[1:])
