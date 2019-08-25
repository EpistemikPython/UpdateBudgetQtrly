##############################################################################################################################
# coding=utf-8
#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year or years
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-13'
__updated__ = '2019-08-25'

from argparse import ArgumentParser
from gnucash import Session
from updateCommon import *
from updateAssets import QTR_SPAN, ASSET_COLS, BASE_YEAR as AST_BY, BASE_YEAR_SPAN as AST_BYS, HDR_SPAN as AST_HS

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


class UpdateBalance:
    def __init__(self, p_filename:str, p_mode:str, p_domain:str, p_debug:bool):
        self.log = SattoLog(p_debug)
        self.log.print_text('UpdateBalance', GREEN)
        self.color = GREY

        # cell(s) with location and value to update on Google sheet
        self.data = []

        self.mode = p_mode
        # Google sheet to update
        self.dest = BAL_2_SHEET
        if '1' in self.mode:
            self.dest = BAL_1_SHEET

        self.gnucash_file = p_filename
        self.domain = p_domain

        self.gncu = GnucashUtilities()
        self.gglu = GoogleUtilities()
        self.util = CommonUtilities()

    BASE_MTHLY_ROW = 19
    BASE_YEAR = 2008
    # number of rows between same quarter in adjacent years
    BASE_YEAR_SPAN = 1
    # number of year groups between header rows
    HDR_SPAN = 9

    def get_data(self) -> list :
        return self.data

    # noinspection PyUnboundLocalVariable
    def fill_today(self):
        """
        Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
        """
        self.log.print_text('UpdateBalance.fill_today()', self.color)
        # calls using 'today' ARE NOT off by one day??
        tdate = today - ONE_DAY
        for item in BALANCE_ACCTS:
            path = BALANCE_ACCTS[item]
            acct_name, acct_sum = self.gncu.get_total_balance(self.root_account, path, tdate, self.currency)

            # need assets NOT INCLUDING house and liabilities, which are reported separately
            if item == HOUSE:
                house_sum = acct_sum
            elif item == LIAB:
                liab_sum = acct_sum
            elif item == ASTS:
                if house_sum is not None and liab_sum is not None:
                    acct_sum = acct_sum - house_sum - liab_sum
                    self.log.print_info("Adjusted assets on {} = ${}".format(today, acct_sum.to_eng_string()), self.color)
                else:
                    self.log.print_error("Do NOT have house sum and liab sum!")

            self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[item], acct_sum, self.data)

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
        self.log.print_text('UpdateBalance.fill_current_year()', self.color)

        for i in range(today.month - 1):
            month_end = date(today.year, i+2, 1)-ONE_DAY
            self.log.print_info("month_end = {}".format(month_end), self.color)

            row = self.BASE_MTHLY_ROW + month_end.month
            # fill LIABS
            acct_name, liab_sum = self.gncu.get_total_balance(self.root_account, BALANCE_ACCTS[LIAB], month_end, self.currency)
            self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum, self.data)

            # fill ASSETS for months NOT covered by the Assets sheet
            if month_end.month % 3 != 0:
                acct_name, acct_sum = self.gncu.get_total_balance(self.root_account, BALANCE_ACCTS[ASTS], month_end, self.currency)
                adjusted_assets = acct_sum - liab_sum
                self.log.print_info("Adjusted assets on {} = ${}".format(month_end, adjusted_assets.to_eng_string()), self.color)
                self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[ASTS], row, adjusted_assets, self.data)
            else:
                self.log.print_info('Update reference to Assets sheet for Mar, June, Sep or Dec', self.color)
                # have to update the CELL REFERENCE to current year/qtr ASSETS
                year_row = BASE_ROW + self.util.year_span(today.year, AST_BY, AST_BYS, AST_HS)
                int_qtr = (month_end.month // 3) - 1
                dest_row = year_row + (int_qtr * QTR_SPAN)
                val_num = '1' if '1' in self.dest else '2'
                value = "='Assets " + val_num + "'!" + ASSET_COLS[TOTAL] + str(dest_row)
                self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[ASTS], row, value, self.data)

            # fill DATE for month column
            self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[MTH], row, str(month_end), self.data)

    def fill_previous_year(self):
        """
        PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
        """
        self.log.print_text('UpdateBalance.fill_current_year()', self.color)

        year = today.year - 1
        for i in range(12-today.month):
            dte = date(year, i+today.month+1, 1)-ONE_DAY
            self.log.print_info("date = {}".format(dte), self.color)

            row = self.BASE_MTHLY_ROW + dte.month
            # fill LIABS
            acct_name, liab_sum = self.gncu.get_total_balance(self.root_account, BALANCE_ACCTS[LIAB], dte, self.currency)
            self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum, self.data)

            # fill ASSETS for months NOT covered by the Assets sheet
            if dte.month % 3 != 0:
                acct_name, acct_sum = self.gncu.get_total_balance(self.root_account, BALANCE_ACCTS[ASTS], dte, self.currency)
                adjusted_assets = acct_sum - liab_sum
                self.log.print_info("Adjusted assets on {} = ${}".format(dte, adjusted_assets.to_eng_string()), self.color)
                self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[ASTS], row, adjusted_assets, self.data)

        # LIABS entry for year end
        year_end = date(year, 12, 31)
        acct_name, liab_sum = self.gncu.get_total_balance(self.root_account, BALANCE_ACCTS[LIAB], year_end, self.currency)
        # month column
        self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[LIAB][MTH], self.BASE_MTHLY_ROW + 12, liab_sum, self.data)
        # year column
        self.fill_year(year)

    def fill_year(self, year:int):
        """
        :param year: to get data for
        """
        year_end = date(year, 12, 31)
        self.log.print_info("UpdateBalance.fill_year(): year_end = {}".format(year_end), self.color)

        # fill LIABS
        acct_name, liab_sum = self.gncu.get_total_balance(self.root_account, BALANCE_ACCTS[LIAB], year_end, self.currency)
        yr_span = self.util.year_span(year, self.BASE_YEAR, self.BASE_YEAR_SPAN, self.HDR_SPAN)
        self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[LIAB][YR], BASE_ROW + yr_span, liab_sum, self.data)

    # noinspection PyAttributeOutsideInit,PyUnboundLocalVariable
    def fill_google_cells(self, p_save):
        """
        Get Balance data for TODAY:
          LIABS, House, FAMILY, XCHALET, TRUST
        OR for the specified year:
          IF CURRENT YEAR: TODAY & LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
          IF PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
          ELSE: LIABS for year
        :param: p_save: save data to json file
        :return: list of cell(s) with location and value to update on Google sheet
        """
        self.log.print_info("UpdateBalance.fill_google_cells()", self.color)

        try:
            gnucash_session = Session(self.gnucash_file, is_new=False)
            self.root_account = gnucash_session.book.get_root_account()
            commod_tab = gnucash_session.book.get_table()
            self.currency = commod_tab.lookup("ISO4217", "CAD")

            if self.domain == 'today':
                self.fill_today()
            elif self.domain == 'allyears':
                self.fill_all_years()
            else:
                year = self.util.get_int_year(self.domain, self.BASE_YEAR)
                if year == today.year:
                    self.fill_current_year()
                elif today.year - year == 1:
                    self.fill_previous_year()
                else:
                    self.fill_year(year)

            # fill update date & time
            self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[DATE], self.BASE_MTHLY_ROW, today.strftime(FILE_DATE_STR), self.data)
            self.gglu.fill_cell(self.dest, BAL_MTHLY_COLS[TODAY], self.BASE_MTHLY_ROW, today.strftime(CELL_TIME_STR), self.data)

            # no save needed, we're just reading...
            gnucash_session.end()

            if len(self.data) > 0:
                fname = "out/updateBalance_{}".format(self.domain)
                self.util.save_to_json(fname, now, self.data)

        except Exception as ge:
            self.log.print_error("Exception: {}!".format(ge))
            if "gnucash_session" in locals() and gnucash_session is not None:
                gnucash_session.end()
            exit(279)

    def send_to_google_sheet(self) -> dict :
        """
        Send the data to the Google sheet
        :return: server response
        """
        self.log.print_info("UpdateBalance.send_to_google_sheet()", self.color)

        if PROD in self.mode:
            return self.gglu.send_data(self.mode, self.data)
        return {'mode':TEST}

# END class UpdateBalance


def process_args():
    arg_parser = ArgumentParser(description='Update the Balance section of my Google Sheet', prog='updateBalance.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST, PROD], help='write to Google sheet or just test')
    required.add_argument('-d', '--domain', required=True, help="'today' | 'current year' | 'previous year' | year | 'allyears'")
    # optional arguments
    arg_parser.add_argument('--json',  action='store_true', help='Write the Google formatted data to a JSON file')
    arg_parser.add_argument('--debug', action='store_true', help='GENERATE DEBUG OUTPUT: MANY LINES!')

    return arg_parser


def process_input_parameters(argv:list):
    args = process_args().parse_args(argv)
    SattoLog.print_text("\nargs = {}".format(args), BROWN)

    if args.debug:
        SattoLog.print_text('Printing ALL Debug output!!', RED)

    if not osp.isfile(args.gnucash_file):
        SattoLog.print_text("File path '{}' does not exist! Exiting...".format(args.gnucash_file), RED)
        exit(282)
    SattoLog.print_text("\nGnucash file = {}".format(args.gnucash_file), CYAN)

    return args.gnucash_file, args.json, args.debug, args.mode, args.domain


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_balance_main(args:list):
    """
    Main: check command line then use UpdateBalance to get the data from Gnucash file and send to Google sheet
    """
    SattoLog.print_text("Parameters = \n{}".format(json.dumps(args, indent=4)), GREEN)
    gnucash_file, save_json, debug, mode, domain = process_input_parameters(args)

    ub_now = dt.now().strftime(DATE_STR_FORMAT)
    SattoLog.print_text("update_balance_main(): Runtime = {}".format(ub_now), BLUE)

    response = 'No Response'
    try:
        updater = UpdateBalance(gnucash_file, mode, domain, debug)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_cells(save_json)

        response = updater.send_to_google_sheet()

        fname = "out/updateBalance_{}-response".format(domain)
        CommonUtilities.save_to_json(fname, ub_now, response)

    except Exception as e:
        msg = "update_balance_main() EXCEPTION!! '{}'".format(repr(e))
        SattoLog.print_warning(msg)

    SattoLog.print_text(" >>> PROGRAM ENDED.\n", GREEN)
    return response


if __name__ == "__main__":
    import sys
    update_balance_main(sys.argv[1:])
