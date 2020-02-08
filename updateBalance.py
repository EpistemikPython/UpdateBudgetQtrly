##############################################################################################################################
# coding=utf-8
#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year or years
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2019-04-13'
__updated__ = '2020-01-27'

from sys import path, argv
from argparse import ArgumentParser
import logging.config as lgconf
from updateAssets import (ASSET_COLS, BASE_YEAR as UA_BASE_YEAR, BASE_YEAR_SPAN as UA_BASE_YEAR_SPAN,
                          HDR_SPAN as UA_HDR_SPAN, QTR_SPAN as UA_QTR_SPAN)
import yaml
path.append('/newdata/dev/git/Python/Gnucash/createGncTxs')
from gnucash_utilities import *
path.append(BASE_PYTHON_FOLDER + 'Google/')
from google_utilities import GoogleUpdate, BASE_ROW

BASE_YEAR:int = 2008
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN:int = 1
# number of year groups between header rows
HDR_SPAN:int = 9

BASE_MTHLY_ROW:int = 24
BASE_TOTAL_WORTH_ROW:int = 25

# path to the accounts in the Gnucash file
BALANCE_ACCTS = {
    HOUSE : [ASTS, HOUSE] ,
    LIAB  : [LIAB]  ,
    TRUST : [TRUST] ,
    CHAL  : [CHAL]  ,
    ASTS  : [ASTS]
}

BAL_MTHLY_COLS = {
    LIAB  : {YR: 'U', MTH: 'L'},
    DATE  : 'E' ,
    TIME  : 'F' ,
    TODAY : 'C' ,
    ASTS  : 'K' ,
    MTH   : 'I'
}

# cell locations in the Google file
BAL_TODAY_RANGES = {
    HOUSE : BASE_TOTAL_WORTH_ROW + 6 ,
    LIAB  : BASE_TOTAL_WORTH_ROW + 8 ,
    TRUST : BASE_TOTAL_WORTH_ROW + 1 ,
    CHAL  : BASE_TOTAL_WORTH_ROW + 2 ,
    ASTS  : BASE_TOTAL_WORTH_ROW + 7
}

# sheet names in Budget Quarterly
BAL_1_SHEET:str = 'Balance 1'
BAL_2_SHEET:str = 'Balance 2'

# load the logging config
with open(YAML_CONFIG_FILE, 'r') as fp:
    log_cfg = yaml.safe_load(fp.read())
lgconf.dictConfig(log_cfg)
lgr = lg.getLogger(LOGGERS[__file__][0])


class UpdateBalance:
    def __init__(self, p_filename:str, p_mode:str, p_domain:str, p_debug:bool):
        self.debug = p_debug
        lgr.info(F"UpdateBalance({p_mode}, {p_domain})")

        self.mode = p_mode
        # Google sheet to update
        self.dest = BAL_2_SHEET
        if '1' in self.mode:
            self.dest = BAL_1_SHEET
        lgr.info(F"dest = {self.dest}")

        self.gnucash_file = p_filename
        self.gnc_session = None
        self.domain = p_domain

        self.gglu = GoogleUpdate(lgr)

    def get_data(self) -> list:
        return self.gglu.get_data()

    def get_balance(self, bal_path, p_date):
        return self.gnc_session.get_total_balance(bal_path, p_date)

    def fill_today(self):
        """
        Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
        """
        lgr.info(get_current_time())
        # calls using 'today' ARE NOT off by one day??
        tdate = now_dt - ONE_DAY
        house_sum = liab_sum = ZERO
        for item in BALANCE_ACCTS:
            acct_sum = self.get_balance(BALANCE_ACCTS[item], tdate)
            # need assets NOT INCLUDING house and liabilities, which are reported separately
            if item == HOUSE:
                house_sum = acct_sum
            elif item == LIAB:
                liab_sum = acct_sum
            elif item == ASTS:
                if house_sum is not None and liab_sum is not None:
                    acct_sum = acct_sum - house_sum - liab_sum
                    lgr.info(F"Adjusted assets on {now_dt} = ${acct_sum.to_eng_string()}")
                else:
                    lgr.info('Do NOT have house sum and liab sum!')

            self.fill_google_cell(BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[item], acct_sum)

    def fill_all_years(self):
        """
        LIABS for all years
        """
        for i in range(now_dt.year - BASE_YEAR - 1):
            year = BASE_YEAR + i
            # fill LIABS
            self.fill_year_end_liabs(year)

    def fill_current_year(self):
        """
        CURRENT YEAR: fill_today() AND: LIABS for ALL completed month_ends; FAMILY for ALL non-3 completed month_ends in year
        """
        self.fill_today()
        lgr.info(get_current_time())

        for i in range(now_dt.month - 1):
            month_end = date(now_dt.year, i + 2, 1) - ONE_DAY
            lgr.debug(F"month_end = {month_end}")

            row = BASE_MTHLY_ROW + month_end.month
            # fill LIABS
            liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], month_end)
            self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)

            # fill ASSETS for months NOT covered by the Assets sheet
            if month_end.month % 3 != 0:
                acct_sum = self.get_balance(BALANCE_ACCTS[ASTS], month_end)
                adjusted_assets = acct_sum - liab_sum
                lgr.debug(F"Adjusted assets on {month_end} = {adjusted_assets.to_eng_string()}")
                self.fill_google_cell(BAL_MTHLY_COLS[ASTS], row, adjusted_assets)
            else:
                lgr.debug('Update reference to Assets sheet for Mar, June, Sep or Dec')
                # have to update the CELL REFERENCE to current year/qtr ASSETS
                year_row = BASE_ROW + year_span(now_dt.year, UA_BASE_YEAR, UA_BASE_YEAR_SPAN, UA_HDR_SPAN)
                int_qtr = (month_end.month // 3) - 1
                dest_row = year_row + (int_qtr * UA_QTR_SPAN)
                val_num = '1' if '1' in self.dest else '2'
                value = "='Assets " + val_num + "'!" + ASSET_COLS[TOTAL] + str(dest_row)
                self.fill_google_cell(BAL_MTHLY_COLS[ASTS], row, value)

            # fill DATE for month column
            self.fill_google_cell(BAL_MTHLY_COLS[MTH], row, str(month_end))

    def fill_previous_year(self):
        """
        PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
        """
        lgr.info(get_current_time())

        year = now_dt.year - 1
        for mth in range(12 - now_dt.month):
            dte = date(year, mth + now_dt.month + 1, 1) - ONE_DAY
            lgr.info(F"date = {dte}")

            row = BASE_MTHLY_ROW + dte.month
            # fill LIABS
            liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], dte)
            self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)

            # fill ASSETS for months NOT covered by the Assets sheet
            if dte.month % 3 != 0:
                acct_sum = self.get_balance(BALANCE_ACCTS[ASTS], dte)
                adjusted_assets = acct_sum - liab_sum
                lgr.info(F"Adjusted assets on {dte} = ${adjusted_assets.to_eng_string()}")
                self.fill_google_cell(BAL_MTHLY_COLS[ASTS], row, adjusted_assets)

            # fill the date in Month column
            self.fill_google_cell(BAL_MTHLY_COLS[MTH], row, str(dte))

        year_end = date(year, 12, 31)
        row = BASE_MTHLY_ROW + 12
        # fill the year-end date in Month column
        self.fill_google_cell(BAL_MTHLY_COLS[MTH], row, str(year_end))

        # LIABS entry for year end
        liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], year_end)
        # month column
        self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)
        # year column
        self.fill_year_end_liabs(year)

    def fill_year_end_liabs(self, year:int):
        """
        :param year: to get data for
        """
        year_end = date(year, 12, 31)
        lgr.info(F"year_end = {year_end}")

        # fill LIABS
        liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], year_end)
        yr_span = year_span(year, BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
        self.fill_google_cell(BAL_MTHLY_COLS[LIAB][YR], BASE_ROW + yr_span, liab_sum)

    def fill_google_cell(self, p_col:str, p_row:int, p_val:str):
        self.gglu.fill_cell(self.dest, p_col, p_row, p_val)

    def fill_google_data(self, p_save:bool):
        """
        Get Balance data for TODAY:
          LIABS, House, FAMILY, XCHALET, TRUST
        OR for the specified year:
          IF CURRENT YEAR: TODAY & LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
          IF PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
          ELSE: LIABS for year
        :param: p_save: save data to json file
        """
        lgr.info(get_current_time())
        try:
            self.gnc_session = GnucashSession(self.mode, self.gnucash_file, BOTH, lgr)
            self.gnc_session.begin_session()

            if self.domain == 'today':
                self.fill_today()
            elif self.domain == 'allyears':
                self.fill_all_years()
            else:
                year = get_int_year(self.domain, BASE_YEAR)
                if year == now_dt.year:
                    self.fill_current_year()
                elif now_dt.year - year == 1:
                    self.fill_previous_year()
                else:
                    self.fill_year_end_liabs(year)

            # record the date & time of this update
            self.fill_google_cell(BAL_MTHLY_COLS[DATE], BASE_MTHLY_ROW, now_dt.strftime(CELL_DATE_STR))
            self.fill_google_cell(BAL_MTHLY_COLS[TIME], BASE_MTHLY_ROW, now_dt.strftime(CELL_TIME_STR))

            # no save needed, we're just reading...
            self.gnc_session.end_session(False)

            if p_save and len(self.get_data()) > 0:
                fname = F"updateBalance_{self.domain}"
                save_to_json(fname, self.get_data())

        except Exception as fgce:
            lgr.error(F"Exception: {repr(fgce)}!")
            if self.gnc_session:
                self.gnc_session.check_end_session(locals())
            raise fgce

# END class UpdateBalance


def process_args() -> ArgumentParser:
    arg_parser = ArgumentParser(description='Update the Balance section of my Google Sheet', prog='updateBalance.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST,SEND+'1',SEND+'2'],
                          help='SEND to Google sheet OR just TEST')
    required.add_argument('-p', '--period', required=True,
                          help=F"'today' | 'current year' | 'previous year' | {BASE_YEAR}..{now_dt.year - 2} | 'allyears'")
    # optional arguments
    arg_parser.add_argument('-l', '--level', type=int, default=lg.INFO, help='set LEVEL of logging output')
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google formatted data to a JSON file')

    return arg_parser


def process_input_parameters(argl:list) -> (str, bool, bool, str, str) :
    args = process_args().parse_args(argl)
    lgr.info(F"\nargs = {args}")

    lgr.info(F"logger level set to {args.level}")

    if not osp.isfile(args.gnucash_file):
        msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
        lgr.warning(msg)
        raise Exception(msg)

    lgr.info(F"\nGnucash file = {args.gnucash_file}")

    return args.gnucash_file, args.ggl_save, args.level, args.mode, args.period


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_balance_main(args:list) -> dict :
    lgr.info(F"Parameters = \n{json.dumps(args, indent=4)}")
    gnucash_file, save_json, level, mode, domain = process_input_parameters(args)

    ub_now = dt.now().strftime(FILE_DATE_FORMAT)

    lgr.setLevel(level)
    lgr.log(level, F"\n\t\tRuntime = {ub_now}")
    debug = lgr.level < lg.INFO

    try:
        updater = UpdateBalance(gnucash_file, mode, domain, debug)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_data(save_json)

        # send data if in PROD mode
        if SEND in mode:
            response = updater.gglu.send_sheets_data()
            fname = F"updateBalance_{domain}-response"
            save_to_json(fname, response, ub_now)
        else:
            response = saved_log_info

    except Exception as be:
        msg = repr(be)
        lgr.error(msg)
        response = {F"update_balance_main() EXCEPTION: {msg}"}

    lgr.info(" >>> PROGRAM ENDED.\n")
    finish_logging(__file__, gnucash_file.split('.')[0], ub_now)
    return response


if __name__ == "__main__":
    update_balance_main(argv[1:])
