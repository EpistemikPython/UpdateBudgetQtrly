##############################################################################################################################
# coding=utf-8
#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Revenue and Expenses
#                     in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2019-03-30'
__updated__ = '2020-03-21'

base_run_file = __file__.split('/')[-1]
print(base_run_file)

from sys import path, argv, exc_info
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import GoogleUpdate, BASE_ROW

BASE_YEAR:int = 2012
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN:int = 11
# number of rows between quarters in the same year
QTR_SPAN:int = 2

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
    """Take data from a Gnucash file and update an Income tab of my Google Budget-Quarterly document"""
    def __init__(self, p_filename:str, p_mode:str, p_lgr:lg.Logger):
        self.gnucash_file = p_filename
        self.gnucash_data = []

        self.gnc_session = None
        self.gglu = GoogleUpdate(p_lgr)

        p_lgr.info(get_current_time())
        self._lgr = p_lgr

        self.mode = p_mode
        # Google sheet to update
        self.all_inc_dest = ALL_INC_2_SHEET
        self.nec_inc_dest = NEC_INC_2_SHEET
        if '1' in self.mode:
            self.all_inc_dest = ALL_INC_SHEET
            self.nec_inc_dest = NEC_INC_SHEET
        p_lgr.debug(F"all_inc_dest = {self.all_inc_dest}")
        p_lgr.debug(F"nec_inc_dest = {self.nec_inc_dest}\n")

    def get_gnucash_data(self) -> list:
        return self.gnucash_data

    def get_google_data(self) -> list:
        return self.gglu.get_data()

    def fill_splits(self, account_path:list, period_starts:list, periods:list):
        self._lgr.log(5, get_current_time())
        return fill_splits(self.gnc_session.get_root_acct(), account_path, period_starts, periods, self._lgr)

    def get_revenue(self, period_starts:list, periods:list) -> dict:
        """
        Get REVENUE data for the specified periods
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :return: revenue for period
        """
        self._lgr.log(5, get_current_time())
        data_quarter = {}
        str_rev = '= '
        for item in REV_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO
            self._lgr.log(5, 'set periods')
            acct_base = REV_ACCTS[item]
            self._lgr.log(5, F"acct_base = {acct_base}")
            acct_name = self.fill_splits(acct_base, period_starts, periods)

            sum_revenue = (periods[0][2] + periods[0][3]) * (-1)
            str_rev += sum_revenue.to_eng_string() + (' + ' if item != EMPL else '')
            self._lgr.debug(F"{acct_name} Revenue for period = ${sum_revenue}")

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
        self._lgr.log(5, get_current_time())
        str_dedns = '= '
        for item in DEDN_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_path = DEDN_ACCTS[item]
            acct_name = self.fill_splits(acct_path, period_starts, periods)

            sum_deductions = periods[0][2] + periods[0][3]
            str_dedns += sum_deductions.to_eng_string() + (' + ' if item != "ML" else '')
            self._lgr.debug(F"{acct_name} {EMPL} Deductions for {p_year}-Q{data_qtr[QTR]} = ${sum_deductions}")

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
        self._lgr.log(5, get_current_time())
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
            self._lgr.debug(F"{acct_name.split('_')[-1]} Expenses for {p_year}-Q{data_qtr[QTR]} = ${str_expenses}")
            str_total += str_expenses + ' + '

        return str_total

    def prepare_gnucash_data(self, save_gnc:bool, p_year:int, p_qtr:int):
        """
        Get REVENUE and EXPENSE data for ONE specified Quarter or ALL four Quarters for the specified Year
        >> NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
           the Gnucash data in a separate dict instead of just preparing a Google data dict
        :param save_gnc: true if want to save the Gnucash data to a JSON file
        :param   p_year: year to update
        :param    p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
        :return: nil
        """
        # get either ONE Quarter or ALL Quarters if updating an entire Year
        num_quarters = 1 if p_qtr else 4
        self._lgr.info("UpdateRevExps.prepare_gnucash_data(): find Revenue & Expenses in {} for {}{}"
                       .format(self.gnucash_file, p_year, ('-Q' + str(p_qtr)) if p_qtr else ''))
        try:
            self.gnc_session = GnucashSession(self.mode, self.gnucash_file, BOTH, self._lgr)
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
                self._lgr.debug(F"\nTOTAL Revenue for {p_year}-Q{qtr} = ${period_list[0][4] * -1}")

                period_list[0][4] = ZERO
                self.get_expenses(period_starts, period_list, p_year, data_quarter)
                self.get_deductions(period_starts, period_list, p_year, data_quarter)
                self._lgr.debug(F"\nTOTAL Expenses for {p_year}-Q{qtr} = {period_list[0][4]}\n")

                self.gnucash_data.append(data_quarter)
                self._lgr.log(5, json.dumps(data_quarter, indent=4))

            # no save needed, we're just reading...
            self.gnc_session.end_session(False)

            if save_gnc:
                fname = F"updateRevExps_gnc-data-{p_year}{('-Q' + str(p_qtr) if p_qtr else '')}"
                self._lgr.info(F"gnucash data file = {save_to_json(fname, self.gnucash_data)}")

        except Exception as fgde:
            fgde_msg = F"prepare_gnucash_data() EXCEPTION: {repr(fgde)}!"
            tb = exc_info()[2]
            self._lgr.error(fgde_msg, tb)
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
        self._lgr.info(get_current_time())
        year_row = BASE_ROW + year_span(p_year, BASE_YEAR, BASE_YEAR_SPAN, 0)
        # get exact row from Quarter value in each item
        for item in self.gnucash_data:
            self._lgr.log(5, F"{QTR} = {item[QTR]}")
            dest_row = year_row + ((get_int_quarter(item[QTR]) - 1) * QTR_SPAN)
            self._lgr.log(5, F"dest_row = {dest_row}\n")
            for key in item:
                if key != QTR:
                    dest = BOOL_NEC_INC
                    if key in (REV, BAL, CONT):
                        dest = BOOL_ALL_INC
                    self.fill_google_cell(dest, REV_EXP_COLS[key], dest_row, item[key])

        # fill update date & time to ALL and NEC
        today_row = BASE_ROW - 1 + year_span(now_dt.year + 2, BASE_YEAR, BASE_YEAR_SPAN, 0)
        self.fill_google_cell(BOOL_NEC_INC, REV_EXP_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(BOOL_NEC_INC, REV_EXP_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))
        self.fill_google_cell(BOOL_ALL_INC, REV_EXP_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(BOOL_ALL_INC, REV_EXP_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))

        str_qtr = None
        if len(self.gnucash_data) == 1:
            str_qtr = self.gnucash_data[0][QTR]

        if save_google:
            fname = F"updateRevExps_google-data-{str(p_year)}{('-Q' + str_qtr if str_qtr else '')}"
            self._lgr.info(F"google data file = {save_to_json(fname, self.get_google_data())}")

# END class UpdateRevExps


def process_args() -> ArgumentParser:
    arg_parser = ArgumentParser(description='Update the Revenues & Expenses section of my Google Sheet', prog='updateRevExps.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST,SEND+'1',SEND+'2'],
                          help='SEND to Google sheet (1 or 2) OR just TEST')
    required.add_argument('-y', '--year', required=True, help=F"year to update: {BASE_YEAR}..2019")
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices=['1','2','3','4'], help="quarter to update: 1..4")
    arg_parser.add_argument('-l', '--level', type=int, default=lg.INFO, help='set LEVEL of logging output')
    arg_parser.add_argument('--gnc_save',  action='store_true', help='Write the Gnucash data to a JSON file')
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google data to a JSON file')
    arg_parser.add_argument('--resp_save', action='store_true', help='Write the Google RESPONSE to a JSON file')

    return arg_parser


def process_input_parameters(argl:list, lgr:lg.Logger) -> (str, bool, bool, bool, str, int, int):
    args = process_args().parse_args(argl)
    # lgr.info(F"\nargs = {args}")

    lgr.info(F"logger level set to {args.level}")

    if not osp.isfile(args.gnucash_file):
        msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
        lgr.error(msg)
        raise Exception(msg)

    lgr.info(F"\n\t\tGnucash file = {args.gnucash_file}")

    year = get_int_year(args.year, BASE_YEAR)
    qtr = 0 if args.quarter is None else get_int_quarter(args.quarter)

    return args.gnucash_file, args.gnc_save, args.ggl_save, args.resp_save, args.level, args.mode, year, qtr


def update_rev_exps_main(args:list) -> dict:
    lgr = get_logger(base_run_file)

    gnucash_file, save_gnc, save_ggl, save_resp, level, mode, target_year, target_qtr = process_input_parameters(args, lgr)

    # get info for log names
    _, fname = osp.split(gnucash_file)
    base_name, _ = osp.splitext(fname)
    target_name = F"-{target_year}{('-Q' + str(target_qtr) if target_qtr else '')}"
    log_name = LOGGERS.get(base_run_file)[1] + '_' + base_name + target_name

    revexp_now = dt.now().strftime(FILE_DATE_FORMAT)

    lgr.setLevel(level)
    lgr.log(level, F"\n\t\tRuntime = {revexp_now}")

    try:
        updater = UpdateRevExps(gnucash_file, mode, lgr)
        # READ the required Gnucash data
        updater.prepare_gnucash_data(save_gnc, target_year, target_qtr)

        # package the Gnucash data in the update format required by Google sheets
        updater.fill_google_data(target_year, save_ggl)

        # check if SENDING data
        if SEND in mode:
            response = updater.gglu.send_sheets_data()
            if save_resp:
                rf_name = F"UpdateRevExps_response{target_name}"
                lgr.info(F"google response file = {save_to_json(rf_name, response, revexp_now)}")
        else:
            response = {'Response':saved_log_info}

    except Exception as reme:
        reme_msg = repr(reme)
        lgr.error(reme_msg)
        response = {'update_rev_exps_main() EXCEPTION':F"{reme_msg}"}

    lgr.info(" >>> PROGRAM ENDED.\n")
    finish_logging(base_run_file, log_name, revexp_now)
    return response

# END class UpdateRevExps


if __name__ == "__main__":
    update_rev_exps_main(argv[1:])
