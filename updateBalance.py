##############################################################################################################################
# coding=utf-8
#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year or years
#
# Copyright (c) 2019-21 Mark Sattolo <epistemik@gmail.com>
#
__author__       = "Mark Sattolo"
__author_email__ = "epistemik@gmail.com"
__created__ = "2019-04-13"
__updated__ = "2021-06-02"

import sys
from updateAssets import ASSETS_DATA, ASSET_COLS
from updateBudget import *

base_run_file = get_base_filename(__file__)
# print(base_run_file)

BALANCE_DATA = {
    # first data row in the sheet
    BASE_ROW  : 4 ,
    # first variable year in google sheet
    BASE_YEAR : 2008 ,
    # NO QUARTER DATA in BALANCE SHEET
    # number of rows to same quarter in adjacent year, NOT INCLUDING header rows
    YEAR_SPAN : 1 ,
    # number of rows to adjacent quarter in the same year
    QTR_SPAN  : 0 ,
    # number of years between header rows
    HDR_SPAN  : 8
}

BASE_MTHLY_ROW:int = 25
BASE_TOTAL_WORTH_ROW:int = 26

# path to the accounts in the Gnucash file
BALANCE_ACCTS = {
    HOUSE   : [FAM, HOUSE] ,
    LIAB    : [FAM, LIAB]  ,
    TRUST   : [TRUST] ,
    CHAL    : [CHAL]  ,
    FAM     : [FAM]   ,
    INVEST  : [FAM, INVEST] ,
    LIQ     : [FAM, LIQ] ,
    PM      : [FAM, PM]  ,
    REW     : [FAM, REW]
}

# column index in the Google sheets
BAL_MTHLY_COLS = {
    LIAB : {YR: 'U', MTH: 'L'},
    DATE : 'E' ,
    TIME : 'F' ,
    TODAY: 'C' ,
    FAM  : 'K' ,
    MTH  : 'I'
}

# cell locations in the Google file
BAL_TODAY_RANGES = {
    HOUSE : BASE_TOTAL_WORTH_ROW + 6 ,
    LIAB  : BASE_TOTAL_WORTH_ROW + 8 ,
    TRUST : BASE_TOTAL_WORTH_ROW + 1 ,
    CHAL  : BASE_TOTAL_WORTH_ROW + 2 ,
    FAM   : BASE_TOTAL_WORTH_ROW + 7
}


class UpdateBalance(UpdateBudget):
    """Take data from a Gnucash file and update a Balance tab of my Google Budget-Quarterly document."""
    def __init__(self, args:list, p_logname:str, p_baseyear:int):
        super().__init__(args, p_logname, p_baseyear)

        # Google sheet to update
        self.dest = BAL_2_SHEET
        if '1' in self.mode:
            self.dest = BAL_1_SHEET
        self._lgr.debug(F"dest = {self.dest}")

        self._gnc_session = None

        # NO saved gnc data for Balance
        self.save_gnc = False

    def get_balance(self, bal_path:list, p_date:date) -> Decimal:
        return self._gnc_session.get_total_balance(bal_path, p_date)

    def fill_today(self):
        """Get Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST."""
        self._lgr.debug(get_current_time())
        # calls using 'today' ARE NOT off by one day??
        tdate = now_dt - ONE_DAY
        asset_sums = {}
        for item in BALANCE_ACCTS:
            acct_sum = self.get_balance(BALANCE_ACCTS[item], tdate)
            if item == HOUSE:
                self.fill_google_cell(BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[item], acct_sum)
            elif item == LIAB:
                self.fill_google_cell(BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[item], acct_sum)
            # need family assets EXCLUDING house and liabilities, which are reported separately
            else:
                asset_sums[item] = acct_sum

        # report the family amount as the sum of the individual accounts
        family_sum = "= " + str(asset_sums[INVEST]) + " + " + str(asset_sums[LIQ]) + " + " \
                     + str(asset_sums[PM]) + " + " + str(asset_sums[REW])
        self._lgr.debug(F"Adjusted assets on {now_dt} = '{family_sum}'")
        self.fill_google_cell(BAL_MTHLY_COLS[TODAY], BAL_TODAY_RANGES[FAM], family_sum)

    def fill_current_year(self):
        """
        CURRENT YEAR: fill_today() AND: LIABS for ALL completed month_ends;
                                        FAMILY for ALL 'non-div-3' completed month_ends in year
        """
        self.fill_today()
        self._lgr.debug(get_current_time())

        for i in range(now_dt.month - 1):
            month_end = date(now_dt.year, i + 2, 1) - ONE_DAY
            self._lgr.debug(F"month_end = {month_end}")

            row = BASE_MTHLY_ROW + month_end.month
            # fill LIABS
            liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], month_end)
            self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)

            # fill ASSETS for months NOT covered by the Assets sheet
            if month_end.month % 3 != 0:
                acct_sum = self.get_balance(BALANCE_ACCTS[FAM], month_end)
                adjusted_assets = acct_sum - liab_sum
                self._lgr.debug(F"Adjusted assets on {month_end} = {adjusted_assets.to_eng_string()}")
                self.fill_google_cell(BAL_MTHLY_COLS[FAM], row, adjusted_assets)
            else:
                self._lgr.debug("Update reference to Assets sheet for Mar, June, Sep or Dec")
                # have to update the CELL REFERENCE to current year/qtr ASSETS
                year_row = ASSETS_DATA[BASE_ROW] \
                           + year_span( now_dt.year, ASSETS_DATA[BASE_YEAR], ASSETS_DATA[YEAR_SPAN], ASSETS_DATA[HDR_SPAN],
                                        self._lgr )
                int_qtr = (month_end.month // 3) - 1
                self._lgr.debug(F"int_qtr = {int_qtr}")
                dest_row = year_row + (int_qtr * ASSETS_DATA.get(QTR_SPAN))
                val_num = '1' if '1' in self.dest else '2'
                value = "='Assets " + val_num + "'!" + ASSET_COLS[TOTAL] + str(dest_row)
                self.fill_google_cell(BAL_MTHLY_COLS[FAM], row, value)

            # fill DATE for month column
            self.fill_google_cell(BAL_MTHLY_COLS[MTH], row, str(month_end))

    def fill_previous_year(self):
        """
        PREVIOUS YEAR: LIABS for ALL NON-completed months;
                       FAMILY assets for ALL 'non-div-3' NON-completed months in year
        """
        self._lgr.debug(get_current_time())

        year = now_dt.year - 1
        for mth in range(12 - now_dt.month):
            dte = date(year, mth + now_dt.month + 1, 1) - ONE_DAY
            self._lgr.debug(F"date = {dte}")

            row = BASE_MTHLY_ROW + dte.month
            # fill LIABS
            liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], dte)
            self.fill_google_cell(BAL_MTHLY_COLS[LIAB][MTH], row, liab_sum)

            # fill ASSETS for months NOT covered by the Assets sheet
            if dte.month % 3 != 0:
                acct_sum = self.get_balance(BALANCE_ACCTS[FAM], dte)
                adjusted_assets = acct_sum - liab_sum
                self._lgr.debug(F"Adjusted assets on {dte} = ${adjusted_assets.to_eng_string()}")
                self.fill_google_cell(BAL_MTHLY_COLS[FAM], row, adjusted_assets)

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
        year_end = date(year, 12, 31)
        self._lgr.debug(F"year_end = {year_end}")

        # fill LIABS
        liab_sum = self.get_balance(BALANCE_ACCTS[LIAB], year_end)
        yr_span = year_span( year, BALANCE_DATA[BASE_YEAR], BALANCE_DATA[YEAR_SPAN], BALANCE_DATA[HDR_SPAN] )
        self.fill_google_cell( BAL_MTHLY_COLS[LIAB][YR], BALANCE_DATA[BASE_ROW] + yr_span, str(liab_sum) )

    # noinspection PyUnusedLocal
    def fill_gnucash_data(self, p_session:GnucashSession, p_qtr:int, p_year:str, data_qtr:dict):
        self._gnc_session = p_session

    def fill_google_cell(self, p_col:str, p_row:int, p_val:FILL_CELL_VAL):
        self._ggl_update.fill_cell(self.dest, p_col, p_row, p_val)

    def fill_google_data(self, p_years:list):
        """
        for each of the specified years:
            IF CURRENT YEAR: TODAY & LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
                             Balance data for TODAY: LIABS, House, FAMILY, XCHALET, TRUST
            IF PREVIOUS YEAR: LIABS for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
        """
        self._lgr.info(F"timespan = {p_years}\n")

        for yr in p_years:
            year = get_int_year( yr, BALANCE_DATA[BASE_YEAR] )
            if year == now_dt.year:
                self.fill_current_year()
            elif now_dt.year - 1 == year:
                self.fill_previous_year()
            else:
                self.fill_year_end_liabs(year)

# END class UpdateBalance


def update_balance_main(args:list) -> dict:
    balance = UpdateBalance(args, base_run_file, BALANCE_DATA[BASE_YEAR])
    return balance.go()


if __name__ == "__main__":
    update_balance_main(sys.argv[1:])
    exit()
