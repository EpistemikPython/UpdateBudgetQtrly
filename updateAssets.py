##############################################################################################################################
# coding=utf-8
#
# updateAssets.py -- use the Gnucash and Google APIs to update the Assets
#                    in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__       = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__created__ = '2019-04-06'
__updated__ = '2020-04-05'

from sys import path, argv
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from updateBudget import *

base_run_file = get_base_filename(__file__)
print(base_run_file)

ASSETS_DATA = {
    # first year row in google sheet
    BASE_YEAR : 2007 ,
    # number of rows between same quarter in adjacent years
    YEAR_SPAN : 4 ,
    # number of rows between quarters in the same year
    QTR_SPAN : 1 ,
    # number of year groups between header rows
    HDR_SPAN : 3
}


class UpdateAssets:
    """Take data from a Gnucash file and update an Assets tab of my Google Budget-Quarterly document"""
    def __init__(self, p_mode:str, p_lgr:lg.Logger):
        self._gnucash_data = []
        self._gglu = GoogleUpdate(p_lgr)

        p_lgr.info(get_current_time())
        self._lgr = p_lgr

        self.mode = p_mode
        # Google sheet to update
        self.dest = QTR_ASTS_2_SHEET
        if '1' in self.mode:
            self.dest = QTR_ASTS_SHEET
        p_lgr.info(F"dest = {self.dest}")

    def get_gnucash_data(self) -> list:
        return self._gnucash_data

    def get_google_data(self) -> list:
        return self._gglu.get_data()

    def fill_gnucash_data(self, p_session:GnucashSession, p_qtr:int, p_year:int, data_qtr:dict) -> dict:
        """
        Get ASSET data for ONE specified Quarter or ALL four Quarters for the specified Year
        :param   p_session: Gnucash session reference
        :param       p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
        :param      p_year: year to update
        :param    data_qtr: dict for data
        """
        self._lgr.info(F"find Assets in {p_session.get_file_name()} for {p_year}{('-Q' + str(p_qtr) if p_qtr else '')}")
        start_month = (p_qtr * 3) - 2
        end_date = current_quarter_end(p_year, start_month)

        p_session.get_account_assets(ASSET_ACCTS, end_date, p_data=data_qtr)
        data_qtr[QTR] = str(p_qtr)

        self._gnucash_data.append(data_qtr)
        return data_qtr

    def fill_google_cell(self, p_col:str, p_row:int, p_val:str):
        self._gglu.fill_cell(self.dest, p_col, p_row, p_val)

    def fill_google_data(self, target_year:int):
        """
        Fill the Google data list.
        :param   target_year: year to update
        """
        self._lgr.info(F"target year = {target_year}\n")
        year_row = BASE_ROW + year_span(target_year, ASSETS_DATA.get(BASE_YEAR),
                                        ASSETS_DATA.get(YEAR_SPAN), ASSETS_DATA.get(HDR_SPAN))
        # get exact row from Quarter value in each item
        for item in self._gnucash_data:
            self._lgr.info(F"{QTR} = {item[QTR]}")
            int_qtr = int(item[QTR])
            dest_row = year_row + ((int_qtr - 1) * ASSETS_DATA.get(QTR_SPAN))
            self._lgr.info(F"dest_row = {dest_row}\n")
            for key in item:
                if key != QTR:
                    # FOR YEAR 2015 OR EARLIER: GET RESP INSTEAD OF Rewards for COLUMN O
                    if key == RESP and target_year > 2015:
                        continue
                    if key == RWRDS and target_year < 2016:
                        continue
                    self.fill_google_cell(ASSET_COLS[key], dest_row, item[key])

    def record_update(self):
        today_row = BASE_ROW + 1 + year_span(now_dt.year + 2, ASSETS_DATA.get(BASE_YEAR),
                                             ASSETS_DATA.get(YEAR_SPAN), ASSETS_DATA.get(HDR_SPAN))
        self.fill_google_cell(ASSET_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
        self.fill_google_cell(ASSET_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))

    def send_sheets_data(self) -> dict:
        return self._gglu.send_sheets_data()

# END class UpdateAssets


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_assets_main(args:list) -> dict:
    updater = UpdateBudget(args, base_run_file, ASSETS_DATA)

    assets = UpdateAssets(updater.get_mode(), updater.get_logger())

    return updater.go(assets)


if __name__ == "__main__":
    update_assets_main(argv[1:])
