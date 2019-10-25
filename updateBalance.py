##############################################################################################################################
# coding=utf-8
#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year or years
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-13'
__updated__ = '2019-10-24'

from sys import path
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
path.append("/home/marksa/dev/git/Python/Google")
from argparse import ArgumentParser
from gnucash_utilities import *
from google_utilities import GoogleUpdate, BASE_ROW
from updateAssets import ASSET_COLS, UpdateAssets as UA

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
    DATE  : 'B' ,
    TODAY : 'C' ,
    ASTS  : 'K' ,
    MTH   : 'I'
}

# cell locations in the Google file
BAL_TODAY_RANGES = {
    HOUSE : 26 ,
    LIAB  : 28 ,
    TRUST : 21 ,
    CHAL  : 22 ,
    ASTS  : 27
}

# sheet names in Budget Quarterly
BAL_1_SHEET:str = 'Balance 1'
BAL_2_SHEET:str = 'Balance 2'


class UpdateBalance:
    def __init__(self, p_filename:str, p_mode:str, p_domain:str, p_debug:bool):
        self.debug = p_debug
        self._logger = SattoLog(my_color=GREY, do_logging=p_debug)
        self._log('UpdateBalance', GREEN)

        self.mode = p_mode
        # Google sheet to update
        self.dest = BAL_2_SHEET
        if '1' in self.mode:
            self.dest = BAL_1_SHEET
        self._log("dest = {}".format(self.dest))

        self.gnucash_file = p_filename
        self.gnc_session = None
        self.domain = p_domain

        self.gglu = GoogleUpdate(self._logger)

    BASE_YEAR:int = 2008
    # number of rows between same quarter in adjacent years
    BASE_YEAR_SPAN:int = 1
    # number of year groups between header rows
    HDR_SPAN:int = 9
    BASE_MTHLY_ROW:int = 19

    def _log(self, p_msg:str, p_color:str=''):
        if self._logger:
            self._logger.print_info(p_msg, p_color, p_frame=inspect.currentframe().f_back)

    def _err(self, p_msg:str, err_frame:FrameType):
        if self._logger:
            self._logger.print_info(p_msg, BR_RED, p_frame=err_frame)

    def get_data(self) -> list:
        return self.gglu.get_data()

    def get_log(self) -> list:
        return self._logger.get_log()

    def get_balance(self, bal_path, p_date):
        return self.gnc_session.get_total_balance(bal_path, p_date)

    def fill_today(self):
        """
        Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
        """
        self._log('UpdateBalance.fill_today()')
        # calls using 'today' ARE NOT off by one day??
        tdate = today - ONE_DAY
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
                    self._log("Adjusted assets on {} = ${}".format(today, acct_sum.to_eng_string()))
                else:
                    self._log("Do NOT have house sum and liab sum!")

            self.fill_google_cell(BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[item], acct_sum)

    def fill_all_years(self):
        """
        LIABS for all years
        """
        for i in range(today.year - self.BASE_YEAR - 1):
            year = self.BASE_YEAR + i
            # fill LIABS
            self.fill_year(year)

    def fill_current_year(self):
        """
        CURRENT YEAR: fill_today() AND: LIABS for ALL completed month_ends; FAMILY for ALL non-3 completed month_ends in year
        """
        self.fill_today()
        self._log('UpdateBalance.fill_current_year()')

        for i in range(today.month - 1):
            month_end = date(today.year, i+2, 1)-ONE_DAY
            self._log("month_end = {}".format(month_end))

            row = self.BASE_MTHLY_ROW + month_end.month
            # fill LIABS
            liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], month_end)
            self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)

            # fill ASSETS for months NOT covered by the Assets sheet
            if month_end.month % 3 != 0:
                acct_sum = self.get_balance(BALANCE_ACCTS[ASTS], month_end)
                adjusted_assets = acct_sum - liab_sum
                self._log(f"Adjusted assets on ${month_end} = ${adjusted_assets.to_eng_string()}")
                self.fill_google_cell(BAL_MTHLY_COLS[ASTS], row, adjusted_assets)
            else:
                self._log('Update reference to Assets sheet for Mar, June, Sep or Dec')
                # have to update the CELL REFERENCE to current year/qtr ASSETS
                year_row = BASE_ROW + year_span(today.year, UA.BASE_YEAR, UA.BASE_YEAR_SPAN, UA.HDR_SPAN)
                int_qtr = (month_end.month // 3) - 1
                dest_row = year_row + (int_qtr * UA.QTR_SPAN)
                val_num = '1' if '1' in self.dest else '2'
                value = "='Assets " + val_num + "'!" + ASSET_COLS[TOTAL] + str(dest_row)
                self.fill_google_cell(BAL_MTHLY_COLS[ASTS], row, value)

            # fill DATE for month column
            self.fill_google_cell(BAL_MTHLY_COLS[MTH], row, str(month_end))

    def fill_previous_year(self):
        """
        PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
        """
        self._log('UpdateBalance.fill_previous_year()')

        year = today.year - 1
        for i in range(12-today.month):
            dte = date(year, i+today.month+1, 1)-ONE_DAY
            self._log("date = {}".format(dte))

            row = self.BASE_MTHLY_ROW + dte.month
            # fill LIABS
            liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], dte)
            self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)

            # fill ASSETS for months NOT covered by the Assets sheet
            if dte.month % 3 != 0:
                acct_sum = self.get_balance(BALANCE_ACCTS[ASTS], dte)
                adjusted_assets = acct_sum - liab_sum
                self._log("Adjusted assets on {} = ${}".format(dte, adjusted_assets.to_eng_string()))
                self.fill_google_cell(BAL_MTHLY_COLS[ASTS], row, adjusted_assets)

        # LIABS entry for year end
        year_end = date(year, 12, 31)
        liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], year_end)
        # month column
        self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], self.BASE_MTHLY_ROW + 12, liab_sum)
        # year column
        self.fill_year(year)

    def fill_year(self, year:int):
        """
        :param year: to get data for
        """
        year_end = date(year, 12, 31)
        self._log("UpdateBalance.fill_year(): year_end = {}".format(year_end))

        # fill LIABS
        liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], year_end)
        yr_span = year_span(year, self.BASE_YEAR, self.BASE_YEAR_SPAN, self.HDR_SPAN)
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
        self._log("UpdateBalance.fill_google_data()")
        try:
            self.gnc_session = GnucashSession(self.mode, self.gnucash_file, self.debug, BOTH)
            self.gnc_session.begin_session()

            if self.domain == 'today':
                self.fill_today()
            elif self.domain == 'allyears':
                self.fill_all_years()
            else:
                year = get_int_year(self.domain, self.BASE_YEAR)
                if year == today.year:
                    self.fill_current_year()
                elif today.year - year == 1:
                    self.fill_previous_year()
                else:
                    self.fill_year(year)

            # record the date & time of this update
            self.fill_google_cell(BAL_MTHLY_COLS[DATE], self.BASE_MTHLY_ROW, today.strftime(FILE_DATE_STR))
            self.fill_google_cell(BAL_MTHLY_COLS[TODAY], self.BASE_MTHLY_ROW, today.strftime(CELL_TIME_STR))

            # no save needed, we're just reading...
            self.gnc_session.end_session(False)

            if p_save and len(self.get_data()) > 0:
                fname = "out/updateBalance_{}".format(self.domain)
                save_to_json(fname, now, self.get_data())

        except Exception as fgce:
            self._err(f"Exception: ${repr(fgce)}!", inspect.currentframe().f_back)
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
                          help="'today' | 'current year' | 'previous year' | {}..{} | 'allyears'"
                               .format(UpdateBalance.BASE_YEAR, today.year - 2))
    # optional arguments
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google formatted data to a JSON file')
    arg_parser.add_argument('--debug', action='store_true', help='GENERATE DEBUG OUTPUT: MANY LINES!')

    return arg_parser


def process_input_parameters(argl:list) -> (str, bool, bool, str, str) :
    args = process_args().parse_args(argl)
    SattoLog.print_text("\nargs = {}".format(args), BROWN)

    if args.debug:
        SattoLog.print_text('Printing ALL Debug output!!', RED)

    if not osp.isfile(args.gnucash_file):
        SattoLog.print_text("File path '{}' does not exist! Exiting...".format(args.gnucash_file), RED)
        exit(291)
    SattoLog.print_text("\nGnucash file = {}".format(args.gnucash_file), CYAN)

    return args.gnucash_file, args.ggl_save, args.debug, args.mode, args.period


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_balance_main(args:list) -> dict :
    SattoLog.print_text("Parameters = \n{}".format(json.dumps(args, indent=4)), GREEN)
    gnucash_file, save_json, debug, mode, domain = process_input_parameters(args)

    ub_now = dt.now().strftime(DATE_STR_FORMAT)
    SattoLog.print_text("update_balance_main(): Runtime = {}".format(ub_now), BLUE)

    try:
        updater = UpdateBalance(gnucash_file, mode, domain, debug)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_data(save_json)

        # send data if in PROD mode
        if SEND in mode:
            response = updater.gglu.send_sheets_data()
            fname = "out/updateBalance_{}-response".format(domain)
            save_to_json(fname, ub_now, response)
        else:
            response = updater.get_log()

    except Exception as be:
        msg = repr(be)
        SattoLog.print_warning(msg)
        response = {"update_balance_main() EXCEPTION:": "{}".format(msg)}

    SattoLog.print_text(" >>> PROGRAM ENDED.\n", GREEN)
    return response


if __name__ == "__main__":
    from sys import argv
    update_balance_main(argv[1:])
