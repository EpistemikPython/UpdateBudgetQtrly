########################################################################################################################################
# coding=utf-8
#
# updateAssets.py -- use the Gnucash and Google APIs to update the Assets in my BudgetQtrly document for a specified period of time
#
# Copyright (c) 2025 Mark Sattolo <epistemik@gmail.com>
#
__author__         = "Mark Sattolo"
__author_email__   = "epistemik@gmail.com"
__python_version__ = "3.10+"
__created__ = "2019-04-06"
__updated__ = "2025-08-01"

from updateBudget import *

ASSETS_DATA = {
    # first data row in the sheet
    BASE_ROW  : 3 ,
    # first data year in google sheet
    BASE_YEAR : 2007 ,
    # number of rows to same quarter in adjacent year, NOT INCLUDING header rows
    YEAR_SPAN : 4 ,
    # number of rows to adjacent quarter in the same year
    QTR_SPAN : 1 ,
    # number of years between header rows
    HDR_SPAN : 3
}

# original paths to accounts in the Gnucash file
ASSET_ACCTS = {
    AU    : [FAM, PM, "Au"],
    AG    : [FAM, PM, "Ag"],
    CASH  : [FAM, LIQ, "$&"],
    BANK  : [FAM, LIQ, BANK],
    REW   : [FAM, REW],
    RESP  : [FAM, INVEST, "xRESP"],
    OPEN  : [FAM, INVEST, OPEN],
    RRSP  : [FAM, INVEST, RRSP],
    TFSA  : [FAM, INVEST, TFSA],
    HOUSE : [FAM, HOUSE]
}

# >> from 2019 onwards
# moved the Loans account to a new position in Gnucash right under Family
# also ignore some of the not very useful sub-accounts in the list of assets,
ASSET_ACCTS_NEW = {
    PM    : [FAM, PM],
    LOAN  : [FAM, LOAN],
    LIQ   : [FAM, LIQ],
    REW   : [FAM, REW], # Rewards
    OPEN  : [FAM, INVEST, OPEN],
    RRSP  : [FAM, INVEST, RRSP],
    TFSA  : [FAM, INVEST, TFSA],
    HOUSE : [FAM, HOUSE]
}

# have the CAR account from 2024 onwards
ASSET_ACCTS_CURRENT = ASSET_ACCTS_NEW | { CAR: [FAM, CAR] }

# column index in the Google sheets
ASSET_COLS = {
    DATE  : 'B',
    AU    : 'U',
    AG    : 'T',
    PM    : 'S',
    CASH  : 'R',
    LOAN  : 'Q',
    BANK  : 'Q',
    LIQ   : 'P',
    REW   : 'O',
    RESP  : 'O',
    OPEN  : 'L',
    RRSP  : 'M',
    TFSA  : 'N',
    HOUSE : 'I',
    TOTAL : 'H',
    CAR   : 'R'
}


class UpdateAssets(UpdateBudget):
    """Take data from a Gnucash file and update an Assets tab of my Google Budget-Quarterly document."""
    def __init__(self, args:list, p_logname:str):
        super().__init__(args, p_logname)

        # Google sheet to update or just testing
        self.dest = self.target
        if '1' in self.target:
            self.dest = QTR_ASTS_SHEET
        elif '2' in self.target:
            self.dest = QTR_ASTS_2_SHEET
        self._lgr.debug(f"dest = {self.dest}")

    def fill_gnucash_data(self, p_session:GnucashSession, p_qtr:int, p_year:str):
        """
        Get ASSET data for specified year and quarter
        :param   p_session: Gnucash session reference
        :param       p_qtr: 1..4 for quarter to update
        :param      p_year: year to update
        """
        self._lgr.debug(f"find Assets in {p_session.get_file_name()} for {p_year}-Q{p_qtr}")
        start_month = (p_qtr * 3) - 2
        int_year = get_int_year( p_year, ASSETS_DATA[BASE_YEAR] )
        end_date = current_quarter_end(int_year, start_month)

        data_qtr = {}
        asset_accounts = ASSET_ACCTS_NEW
        # had slightly different accounts before 2019
        if int_year < 2019:
            asset_accounts = ASSET_ACCTS
        elif int_year > 2023:
            asset_accounts = ASSET_ACCTS_CURRENT
        p_session.get_account_assets(asset_accounts, end_date, p_data=data_qtr)
        data_qtr[YR] = p_year
        data_qtr[QTR] = str(p_qtr)

        self._gnucash_data.append(data_qtr)
        self._lgr.debug(json.dumps(data_qtr, indent = 4))

    def fill_google_data(self, p_years:list):
        """
        Fill the Google data list.
        :param p_years: timespan to update
        """
        self._lgr.info(f"timespan = {p_years}\n")
        # get the row from Year and Quarter value in each item
        for item in self._gnucash_data:
            target_year = get_int_year( item[YR], ASSETS_DATA[BASE_YEAR] )
            year_row = ASSETS_DATA[BASE_ROW] + year_span( target_year, ASSETS_DATA[BASE_YEAR], ASSETS_DATA[YEAR_SPAN], ASSETS_DATA[HDR_SPAN] )
            int_qtr = int(item[QTR])
            dest_row = year_row + ( (int_qtr - 1) * ASSETS_DATA[QTR_SPAN] )
            self._lgr.info(f"{item[YR]}-Q{item[QTR]} dest row = {dest_row}\n")
            for key in item:
                if key not in (QTR,YR):
                    # FOR YEAR 2015 OR EARLIER: GET RESP INSTEAD OF Rewards for COLUMN O
                    if key == RESP and target_year > 2015:
                        continue
                    if key == REW and target_year < 2016:
                        continue
                    self._ggl_update.fill_cell(self.dest, ASSET_COLS[key], dest_row, item[key])
# END class UpdateAssets


def update_assets_main(args:list) -> dict:
    assets = UpdateAssets(args, get_base_filename(__file__))
    return assets.go("Assets")


if __name__ == "__main__":
    update_assets_main(argv[1:])
    exit()
