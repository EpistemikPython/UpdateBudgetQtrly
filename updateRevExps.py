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
__updated__ = '2020-03-29'

from sys import argv
from updateBudget import *

base_run_file = get_base_filename(__file__)
print(base_run_file)

BASE_YEAR:int = 2012
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN:int = 11
# number of rows between quarters in the same year
QTR_SPAN:int = 2


class UpdateRevExps:
    """Take data from a Gnucash file and update an Income tab of my Google Budget-Quarterly document"""
    def __init__(self, p_mode:str, p_lgr:lg.Logger):
        self.gnucash_data = []
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

    def fill_splits(self, root_acct:Account, account_path:list, period_starts:list, periods:list) -> str:
        self._lgr.log(5, get_current_time())
        if root_acct:
            return fill_splits(root_acct, account_path, period_starts, periods, self._lgr)
        self._lgr.error("NO root account!")
        return ''

    def get_data(self, p_root_acct:Account, p_qtr:int, period_starts:list, periods:list, p_year:int, data_qtr:dict) -> dict:
        self.get_revenue(p_root_acct, period_starts, periods, data_qtr)
        data_qtr[QTR] = str(p_qtr)
        self._lgr.debug(F"\nTOTAL Revenue for {p_year}-Q{p_qtr} = ${periods[0][4] * -1}")

        periods[0][4] = ZERO
        self.get_expenses(p_root_acct, period_starts, periods, p_year, data_qtr)
        self.get_deductions(p_root_acct, period_starts, periods, p_year, data_qtr)

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
        self._lgr.log(5, get_current_time())
        str_rev = '= '
        for item in REV_ACCTS:
            # reset the debit and credit totals for each individual account
            periods[0][2] = ZERO
            periods[0][3] = ZERO
            self._lgr.log(5, 'set periods')
            acct_base = REV_ACCTS[item]
            self._lgr.log(5, F"acct_base = {acct_base}")
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
        self._lgr.log(5, get_current_time())
        str_dedns = '= '
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
        self._lgr.log(5, get_current_time())
        str_total = ''
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

    def fill_google_cell(self, p_dest:str, p_col:str, p_row:int, p_time:str):
        self.gglu.fill_cell(p_dest, p_col, p_row, p_time)

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
                    dest = self.nec_inc_dest
                    if key in (REV, BAL, CONT):
                        dest = self.all_inc_dest
                    self.fill_google_cell(dest, REV_EXP_COLS[key], dest_row, item[key])

        self.fill_update_date_and_time()

        str_qtr = None
        if len(self.gnucash_data) == 1:
            str_qtr = self.gnucash_data[0][QTR]

        if save_google:
            fname = F"updateRevExps_google-data-{str(p_year)}{('-Q' + str_qtr if str_qtr else '')}"
            self._lgr.info(F"google data file = {save_to_json(fname, self.get_google_data())}")

    # TODO: keep record of all changes: what exactly and when
    def fill_update_date_and_time(self):
        """fill update date & time to ALL and NEC"""
        today_row = BASE_ROW - 1 + year_span(now_dt.year + 2, BASE_YEAR, BASE_YEAR_SPAN, 0)
        self.fill_google_cell(self.nec_inc_dest, REV_EXP_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(self.nec_inc_dest, REV_EXP_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))
        self.fill_google_cell(self.all_inc_dest, REV_EXP_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(self.all_inc_dest, REV_EXP_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))

# END class UpdateRevExps


def update_re_main(args:list) -> dict:
    lgr = get_logger(base_run_file)

    updater = UpdateBudget(args, base_run_file, BASE_YEAR, BASE_YEAR_SPAN, QTR_SPAN, lgr)

    rev_exp = UpdateRevExps(updater.get_mode(), lgr)

    return updater.go(rev_exp)


def update_rev_exps_main(args:list) -> dict:
    return update_re_main(args)

    # lgr = get_logger(base_run_file)
    #
    # gnucash_file, save_gnc, save_ggl, save_resp, level, mode, target_year, target_qtr = process_input_parameters(args, lgr)
    #
    # # get info for log names
    # _, fname = osp.split(gnucash_file)
    # base_name, _ = osp.splitext(fname)
    # target_name = F"-{target_year}{('-Q' + str(target_qtr) if target_qtr else '')}"
    # log_name = get_logger_filename(base_run_file) + '_' + base_name + target_name
    #
    # revexp_now = dt.now().strftime(FILE_DATE_FORMAT)
    #
    # lgr.setLevel(level)
    # lgr.log(level, F"\n\t\tRuntime = {revexp_now}")
    #
    # try:
    #     updater = UpdateRevExps(gnucash_file, mode, lgr)
    #     # READ the required Gnucash data
    #     updater.prepare_gnucash_data(save_gnc, target_year, target_qtr)
    #
    #     # package the Gnucash data in the update format required by Google sheets
    #     updater.fill_google_data(target_year, save_ggl)
    #
    #     # check if SENDING data
    #     if SEND in mode:
    #         response = updater.gglu.send_sheets_data()
    #         if save_resp:
    #             rf_name = F"UpdateRevExps_response{target_name}"
    #             lgr.info(F"google response file = {save_to_json(rf_name, response, revexp_now)}")
    #     else:
    #         response = {'Response':saved_log_info}
    #
    # except Exception as reme:
    #     reme_msg = repr(reme)
    #     lgr.error(reme_msg)
    #     response = {'update_rev_exps_main() EXCEPTION':F"{reme_msg}"}
    #
    # lgr.info(" >>> PROGRAM ENDED.\n")
    # finish_logging(base_run_file, log_name, revexp_now)
    # return response

# END class UpdateRevExps


if __name__ == "__main__":
    update_re_main(argv[1:])
