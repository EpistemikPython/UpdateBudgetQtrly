##############################################################################################################################
# coding=utf-8
#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Revenue and Expenses
#                     in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-03-30'
__updated__ = '2019-09-06'

from sys import path
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
path.append("/home/marksa/dev/git/Python/Google")
from argparse import ArgumentParser
from gnucash_utilities import *
from google_utilities import *
from investment import *

# path to the account in the Gnucash file
REV_ACCTS = {
    INV : ["REV_Invest"],
    OTH : ["REV_Other"],
    EMP : ["REV_Employment"]
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

BOOL_NEC_INC = False
BOOL_ALL_INC = True


class UpdateRevExps:
    def __init__(self, p_filename:str, p_mode:str, p_debug:bool):
        self.log = SattoLog(p_debug)
        self.log.print_info('UpdateRevExps', GREEN)
        self.color = BROWN

        self.gnucash_file = p_filename
        self.gnucash_data = []
        # cell(s) with location and value to update on Google sheet
        self.google_data = []

        self.mode = p_mode
        # Google sheet to update
        self.all_inc_dest = ALL_INC_2_SHEET
        self.nec_inc_dest = NEC_INC_2_SHEET
        if '1' in self.mode:
            self.all_inc_dest = ALL_INC_SHEET
            self.nec_inc_dest = NEC_INC_SHEET
        self.log.print_info("all_inc_dest = {}".format(self.all_inc_dest), self.color)
        self.log.print_info("nec_inc_dest = {}\n".format(self.nec_inc_dest), self.color)

    BASE_YEAR:int = 2012
    # number of rows between same quarter in adjacent years
    BASE_YEAR_SPAN:int = 11
    # number of rows between quarters in the same year
    QTR_SPAN:int = 2

    def get_gnucash_data(self) -> list:
        return self.gnucash_data

    def get_google_data(self) -> list:
        return self.google_data

    def get_log(self) -> list :
        return self.log.get_log()

    def get_revenue(self, root_account:Account, period_starts:list, periods:list, p_year:int, p_qtr:int) -> dict :
        """
        Get REVENUE data for the specified Quarter
        :param  root_account: Gnucash Account from the Gnucash book
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param         p_qtr: quarter to read: 1..4
        :return: revenue for period
        """
        self.log.print_info('UpdateRevExps.get_revenue()', self.color)
        data_quarter = {}
        str_rev = '= '
        for item in REV_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_base = REV_ACCTS[item]
            acct_name = fill_splits(root_account, acct_base, period_starts, periods)

            sum_revenue = (periods[0][2] + periods[0][3]) * (-1)
            str_rev += sum_revenue.to_eng_string() + (' + ' if item != EMP else '')
            self.log.print_info("{} Revenue for {}-Q{} = ${}".format(acct_name, p_year, p_qtr, sum_revenue), self.color)

        data_quarter[REV] = str_rev
        return data_quarter

    def get_deductions(self, root_account:Account, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> str :
        """
        Get SALARY DEDUCTIONS data for the specified Quarter
        :param  root_account: Gnucash Account from the Gnucash book
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param      data_qtr: collected data for the quarter
        :return: deductions for period
        """
        self.log.print_info('UpdateRevExps.get_deductions()', self.color)
        str_dedns = '= '
        for item in DEDN_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_path = DEDN_ACCTS[item]
            acct_name = fill_splits(root_account, acct_path, period_starts, periods)

            sum_deductions = periods[0][2] + periods[0][3]
            str_dedns += sum_deductions.to_eng_string() + (' + ' if item != "ML" else '')
            self.log.print_info("{} {} Deductions for {}-Q{} = ${}"
                                .format(acct_name, EMP, p_year, data_qtr[QTR], sum_deductions), self.color)

        data_qtr[DEDNS] = str_dedns
        return str_dedns

    def get_expenses(self, root_account:Account, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> str :
        """
        Get EXPENSE data for the specified Quarter
        :param  root_account: Gnucash Account from the Gnucash book
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param      data_qtr: collected data for the quarter
        :return: total expenses for period
        """
        self.log.print_info('UpdateRevExps.get_expenses()', self.color)
        str_total = ''
        for item in EXP_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_base = EXP_ACCTS[item]
            acct_name = fill_splits(root_account, acct_base, period_starts, periods)

            sum_expenses = periods[0][2] + periods[0][3]
            str_expenses = sum_expenses.to_eng_string()
            data_qtr[item] = str_expenses
            self.log.print_info("{} Expenses for {}-Q{} = ${}"
                                .format(acct_name.split('_')[-1], p_year, data_qtr[QTR], str_expenses), self.color)
            str_total += str_expenses + ' + '

        self.get_deductions(root_account, period_starts, periods, p_year, data_qtr)

        return str_total

    # noinspection PyUnboundLocalVariable
    def fill_gnucash_data(self, save_gnc:bool, p_year:int, p_qtr:int):
        """
        Get REVENUE and EXPENSE data for ONE specified Quarter or ALL four Quarters for the specified Year
        >> NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
           the Gnucash data in a separate dict instead of just preparing a Google data dict
        :param save_gnc: save the Gnucash data to a JSON file
        :param   p_year: year to update
        :param    p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
        """
        num_quarters = 1 if p_qtr else 4
        self.log.print_info("URE.fill_gnucash_data(): find Revenue & Expenses in {} for {}{}"
                            .format(self.gnucash_file, p_year, ('-Q' + str(p_qtr)) if p_qtr else ''), self.color)
        try:
            gnucash_session = Session(self.gnucash_file, is_new=False)
            root_account = gnucash_session.book.get_root_account()
            # self.log.print_info("type root_account = {}".format(type(root_account)))

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

                data_quarter = self.get_revenue(root_account, period_starts, period_list, p_year, qtr)
                data_quarter[QTR] = str(qtr)
                self.log.print_info("\n{} Revenue for {}-Q{} = ${}"
                                    .format("TOTAL", p_year, qtr, period_list[0][4] * (-1)), self.color)

                period_list[0][4] = ZERO
                self.get_expenses(root_account, period_starts, period_list, p_year, data_quarter)
                self.log.print_info("\n{} Expenses for {}-Q{} = ${}\n"
                                    .format("TOTAL", p_year, qtr, period_list[0][4]), self.color)

                self.gnucash_data.append(data_quarter)
                self.log.print_info(json.dumps(data_quarter, indent=4), self.color)

            # no save needed, we're just reading...
            gnucash_session.end()

            fname = "out/updateRevExps_gnc-data-{}{}".format(p_year, ('-Q' + str(p_qtr)) if p_qtr else '')
            save_to_json(fname, now, self.gnucash_data)

        except Exception as ge:
            self.log.print_error("Exception: {}!".format(repr(ge)))
            if "gnucash_session" in locals() and gnucash_session is not None:
                gnucash_session.end()
            raise ge

    def fill_google_cell(self, p_all:bool, p_col:str, p_row:int, p_time:str):
        dest = self.nec_inc_dest
        if p_all:
            dest = self.all_inc_dest
        fill_cell(dest, p_col, p_row, p_time, self.google_data)

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
        """
        self.log.print_info('UpdateRevExps.fill_google_data()', self.color)
        year_row = BASE_ROW + year_span(p_year, self.BASE_YEAR, self.BASE_YEAR_SPAN, 0)
        # get exact row from Quarter value in each item
        for item in self.gnucash_data:
            self.log.print_info("{} = {}".format(QTR, item[QTR]))
            dest_row = year_row + ((get_int_quarter(item[QTR]) - 1) * self.QTR_SPAN)
            self.log.print_info("dest_row = {}\n".format(dest_row), self.color)
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
            save_to_json(fname, now, self.google_data)

# END class UpdateRevExps


def process_args() -> ArgumentParser:
    arg_parser = ArgumentParser(description='Update the Revenues & Expenses section of my Google Sheet', prog='updateRevExps.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST, PROD+'1', PROD+'2'],
                          help='write to Google sheet (1 or 2) OR just test')
    required.add_argument('-y', '--year', required=True, help="year to update: {}..2019".format(UpdateRevExps.BASE_YEAR))
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices=['1','2','3','4'], help="quarter to update: 1..4")
    arg_parser.add_argument('--gnc_save',  action='store_true', help='Write the Gnucash formatted data to a JSON file')
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google formatted data to a JSON file')
    arg_parser.add_argument('--debug', action='store_true', help='GENERATE DEBUG OUTPUT: MANY LINES!')

    return arg_parser


def process_input_parameters(argl:list) -> (str, bool, bool, bool, str, int, int) :
    args = process_args().parse_args(argl)
    SattoLog.print_text("\nargs = {}".format(args), BROWN)

    if args.debug:
        SattoLog.print_text('Printing ALL Debug output!!', RED)

    if not osp.isfile(args.gnucash_file):
        SattoLog.print_text("File path '{}' does not exist! Exiting...".format(args.gnucash_file), RED)
        exit(318)
    SattoLog.print_text("\nGnucash file = {}".format(args.gnucash_file), GREEN)

    year = get_int_year(args.year, UpdateRevExps.BASE_YEAR)
    qtr = 0 if args.quarter is None else get_int_quarter(args.quarter)

    return args.gnucash_file, args.gnc_save, args.ggl_save, args.debug, args.mode, year, qtr


def update_rev_exps_main(args:list) -> dict :
    SattoLog.print_text("Parameters = \n{}".format(json.dumps(args, indent=4)), BROWN)
    gnucash_file, save_gnc, save_ggl, debug, mode, target_year, target_qtr = process_input_parameters(args)

    revexp_now = dt.now().strftime(DATE_STR_FORMAT)
    SattoLog.print_text("update_rev_exps_main(): Runtime = {}".format(revexp_now), BLUE)

    try:
        updater = UpdateRevExps(gnucash_file, mode, debug)

        # either for One Quarter or for Four Quarters if updating an entire Year
        updater.fill_gnucash_data(save_gnc, target_year, target_qtr)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_data(target_year, save_ggl)

        # send data if in PROD mode
        if PROD in mode:
            response = send_sheets_data(updater.get_google_data())
            fname = "out/updateRevExps_response-{}{}".format(target_year , ('-Q' + str(target_qtr)) if target_qtr else '')
            save_to_json(fname, revexp_now, response)
        else:
            response = updater.get_log()

    except Exception as ree:
        msg = repr(ree)
        SattoLog.print_warning(msg)
        response = {"update_rev_exps_main() EXCEPTION:": "{}".format(msg)}

    SattoLog.print_text(" >>> PROGRAM ENDED.\n", GREEN)
    return response


if __name__ == "__main__":
    from sys import argv
    update_rev_exps_main(argv[1:])
