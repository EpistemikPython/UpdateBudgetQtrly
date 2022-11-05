##############################################################################################################################
# coding=utf-8
#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Revenue and Expenses
#                     in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2019-21 Mark Sattolo <epistemik@gmail.com>
#
__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2019-03-30'
__updated__ = '2021-07-10'

from updateBudget import *

base_run_file = get_base_filename(__file__)
# print(base_run_file)

REVEXPS_DATA = {
    # first data row in the sheet
    BASE_ROW  : 3 ,
    # first data year in google sheet
    BASE_YEAR : 2008 ,
    # number of rows to same quarter in adjacent year, NOT INCLUDING header rows
    YEAR_SPAN : 10 ,
    # number of rows to adjacent quarter in the same year
    QTR_SPAN  : 2 ,
    # number of years between header rows
    HDR_SPAN  : 1
}

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
DEDNS_BASE = 'DEDNS_Income'
DEDN_ACCTS = {
    "Mark" : [DEDNS_BASE, 'Mark'],
    "Lulu" : [DEDNS_BASE, 'Lulu'],
    "ML"   : [DEDNS_BASE, 'Marie-Laure']
}

# column index in the Google sheets
REV_EXP_COLS = {
    DATE  : 'B', # both
    REV   : 'D', # All Inc
    BAL   : 'P', # All Inc
    CONT  : 'O', # All Inc
    NEC   : 'G', # Nec Inc
    DEDNS : 'D'  # Nec Inc
}


class UpdateRevExps(UpdateBudget):
    """Take data from a Gnucash file and update an Income tab of my Google Budget-Quarterly document."""
    def __init__(self, args:list, p_logname:str):
        super().__init__(args, p_logname)

        # Google sheet to update
        self.all_inc_dest = ALL_INC_2_SHEET
        self.nec_inc_dest = NEC_INC_2_SHEET
        if '1' in self.mode:
            self.all_inc_dest = ALL_INC_SHEET
            self.nec_inc_dest = NEC_INC_SHEET
        self._lgr.debug(F"all_inc_dest = {self.all_inc_dest}")
        self._lgr.debug(F"nec_inc_dest = {self.nec_inc_dest}\n")

    def fill_splits(self, root_acct:Account, account_path:list, period_starts:list, periods:list) -> str:
        self._lgr.debug(get_current_time())
        if root_acct:
            return fill_splits(root_acct, account_path, period_starts, periods, self._lgr)
        self._lgr.error("NO root account!")
        return ""

    def fill_gnucash_data(self, p_session:GnucashSession, p_qtr:int, p_year:str, data_qtr:dict) -> dict:
        root_acct = p_session.get_root_acct()
        start_month = (p_qtr * 3) - 2
        int_year = get_int_year( p_year, REVEXPS_DATA[BASE_YEAR] )

        # for each period keep the start date, end date, debit and credit sums & overall total
        period_list = [
            [
                start_date, end_date,
                ZERO,  # debits sum
                ZERO,  # credits sum
                ZERO  # TOTAL
            ]
            for start_date, end_date in generate_quarter_boundaries(int_year, start_month, 1)
        ]
        # a copy of the above list with just the period start dates
        period_starts = [e[0] for e in period_list]

        self.get_revenue(root_acct, period_starts, period_list, data_qtr)
        data_qtr[YR] = p_year
        data_qtr[QTR] = str(p_qtr)
        self._lgr.debug(F"\n\t\tTOTAL Revenue for {p_year}-Q{p_qtr} = ${period_list[0][4] * -1}")

        period_list[0][4] = ZERO
        self.get_expenses(root_acct, period_starts, period_list, int_year, data_qtr)
        self._lgr.debug(F"\n\t\tTOTAL Expenses for {p_year}-Q{p_qtr} = {period_list[0][4]}\n")

        self.get_deductions(root_acct, period_starts, period_list, int_year, data_qtr)

        self._gnucash_data.append(data_qtr)
        self._lgr.debug(json.dumps(data_qtr, indent = 4))

        return data_qtr

    def get_revenue(self, root_acct:Account, period_starts:list, periods:list, data_qtr:dict) -> str:
        """
        Get REVENUE data for the specified periods
        :param     root_acct: in Gnucash file
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param      data_qtr: dict for data for the specified quarter
        :return: revenue for period
        """
        self._lgr.debug(get_current_time())
        str_rev = "= "
        for item in REV_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO
            self._lgr.debug('set periods')
            acct_base = REV_ACCTS[item]
            self._lgr.debug(F"acct_base = {acct_base}")
            acct_name = self.fill_splits(root_acct, acct_base, period_starts, periods)

            sum_revenue = (periods[0][2] + periods[0][3]) * (-1)
            str_rev += sum_revenue.to_eng_string() + (' + ' if item != EMPL else '')
            self._lgr.debug(F"{acct_name} Revenue for period = ${sum_revenue}")

        data_qtr[REV] = str_rev
        return str_rev

    def get_deductions(self, root_acct:Account, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> str:
        """
        Get SALARY DEDUCTIONS data for the specified Quarter
        :param     root_acct: in Gnucash file
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param      data_qtr: dict for data for the specified quarter
        :return: deductions for period
        """
        self._lgr.debug(get_current_time())
        str_dedns = "= "
        for item in DEDN_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_path = DEDN_ACCTS[item]
            acct_name = self.fill_splits(root_acct, acct_path, period_starts, periods)

            sum_deductions = periods[0][2] + periods[0][3]
            str_dedns += sum_deductions.to_eng_string() + (' + ' if item != "ML" else '')
            self._lgr.debug(F"{acct_name} {EMPL} Deductions for {p_year}-Q{data_qtr[QTR]} = ${sum_deductions}")

        data_qtr[DEDNS] = str_dedns
        return str_dedns

    def get_expenses(self, root_acct:Account, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> str:
        """
        Get EXPENSE data for the specified Quarter
        :param     root_acct: in Gnucash file
        :param period_starts: start date for each period
        :param       periods: structs with the dates and amounts for each quarter
        :param        p_year: year to read
        :param      data_qtr: dict for data for the specified quarter
        :return: total expenses for period
        """
        self._lgr.debug(get_current_time())
        str_total = ""
        for item in EXP_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO

            acct_base = EXP_ACCTS[item]
            acct_name = self.fill_splits(root_acct, acct_base, period_starts, periods)

            sum_expenses = periods[0][2] + periods[0][3]
            str_expenses = sum_expenses.to_eng_string()
            data_qtr[item] = str_expenses
            self._lgr.debug(F"{acct_name.split('_')[-1]} Expenses for {p_year}-Q{data_qtr[QTR]} = ${str_expenses}")
            str_total += str_expenses + ' + '

        return str_total

    def fill_google_data(self, p_years:list):
        """
        Fill the data list:
        for each item in the gnucash data:
            create 5 cells, one each for REV, BAL, CONT, NEC, DEDNS:
            fill in the range based on the year and quarter
            range = SHEET_NAME + '!' + calculated cell
            fill in the values based on the sheet being updated and the type of cell data
            REV string is '= ${INV} + ${OTH} + ${SAL}'
            DEDNS string is '= ${Mk-Dedns} + ${Lu-Dedns} + ${ML-Dedns}'
            others are just the string from the item
        :param p_years: timespan to update
        """
        self._lgr.info(F"timespan = {p_years}\n")
        # get the row from Year and Quarter value in each item
        for item in self._gnucash_data:
            self._lgr.debug(F"item = {item}")
            target_year = get_int_year( item[YR], REVEXPS_DATA[BASE_YEAR] )
            year_row = REVEXPS_DATA[BASE_ROW]\
                       + year_span(target_year, REVEXPS_DATA[BASE_YEAR], REVEXPS_DATA[YEAR_SPAN], REVEXPS_DATA[HDR_SPAN], self._lgr)
            dest_row = year_row + ( (get_int_quarter(item[QTR]) - 1) * REVEXPS_DATA[QTR_SPAN] )
            self._lgr.debug(F"{item[YR]}-Q{item[QTR]} dest row = {dest_row}\n")
            for key in item:
                if key not in (YR,QTR):
                    dest = self.nec_inc_dest
                    if key in (REV, BAL, CONT):
                        dest = self.all_inc_dest
                    self._ggl_update.fill_cell(dest, REV_EXP_COLS[key], dest_row, item[key])

# END class UpdateRevExps


def update_rev_exps_main(args:list) -> dict:
    rev_exp = UpdateRevExps(args, base_run_file)
    return rev_exp.go()


if __name__ == "__main__":
    update_rev_exps_main(argv[1:])
    exit()
