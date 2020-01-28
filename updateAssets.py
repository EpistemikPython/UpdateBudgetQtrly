##############################################################################################################################
# coding=utf-8
#
# updateAssets.py -- use the Gnucash and Google APIs to update the Assets
#                    in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__  = 3.9
__gnucash_version__ = 3.8
__created__ = '2019-04-06'
__updated__ = '2020-01-27'

from sys import path, argv
import yaml
import logging.config as lgconf
from argparse import ArgumentParser
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
from gnucash_utilities import *
path.append("/home/marksa/dev/git/Python/Google")
from google_utilities import GoogleUpdate, BASE_ROW

BASE_YEAR:int = 2007
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN:int = 4
# number of rows between quarters in the same year
QTR_SPAN:int = 1
# number of year groups between header rows
HDR_SPAN:int = 3

# path to the account in the Gnucash file
ASSET_ACCTS = {
    AU    : ["FAMILY", "Prec Metals", "Au"],
    AG    : ["FAMILY", "Prec Metals", "Ag"],
    CASH  : ["FAMILY", "LIQUID", "$&"],
    BANK  : ["FAMILY", "LIQUID", BANK],
    RWRDS : ["FAMILY", RWRDS],
    RESP  : ["FAMILY", "INVEST", "xRESP"],
    OPEN  : ["FAMILY", "INVEST", OPEN],
    RRSP  : ["FAMILY", "INVEST", RRSP],
    TFSA  : ["FAMILY", "INVEST", TFSA],
    HOUSE : ["FAMILY", HOUSE]
}

# column index in the Google sheets
ASSET_COLS = {
    DATE  : 'B',
    AU    : 'U',
    AG    : 'T',
    CASH  : 'R',
    BANK  : 'Q',
    RWRDS : 'O',
    RESP  : 'O',
    OPEN  : 'L',
    RRSP  : 'M',
    TFSA  : 'N',
    HOUSE : 'I',
    TOTAL : 'H'
}

# sheet names in Budget Quarterly
QTR_ASTS_SHEET:str   = 'Assets 1'
QTR_ASTS_2_SHEET:str = 'Assets 2'

MULTI_LOGGING = True
print('MULTI_LOGGING = True')
# load the logging config
with open(YAML_CONFIG_FILE, 'r') as fp:
    log_cfg = yaml.safe_load(fp.read())
lgconf.dictConfig(log_cfg)
lgr = lg.getLogger('gnucash')


class UpdateAssets:
    def __init__(self, p_filename:str, p_mode:str, p_debug:bool):
        self.debug = p_debug
        #
        lgr.info('UpdateAssets')

        self.gnucash_file = p_filename
        self.gnucash_data = []

        self.ggl_update = GoogleUpdate(lgr)

        self.mode = p_mode
        # Google sheet to update
        self.dest = QTR_ASTS_2_SHEET
        if '1' in self.mode:
            self.dest = QTR_ASTS_SHEET
        lgr.info(F"dest = {self.dest}")

    def get_gnucash_data(self) -> list:
        return self.gnucash_data

    def get_google_data(self) -> list:
        return self.ggl_update.get_data()

    def prepare_gnucash_data(self, save_gnc:bool, p_year:int, p_qtr:int):
        """
        Get ASSET data for ONE specified Quarter or ALL four Quarters for the specified Year
        :param save_gnc: save gnucash data to json file
        :param   p_year: year to update
        :param    p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
        """
        lgr.info("UpdateAssets.prepare_gnucash_data(): find Assets in {} for {}{}"
                  .format(self.gnucash_file, p_year, ('-Q' + str(p_qtr)) if p_qtr else ''))
        num_quarters = 1 if p_qtr else 4
        gnc_session = None
        try:
            gnc_session = GnucashSession(self.mode, self.gnucash_file, BOTH, lgr)
            gnc_session.begin_session()

            for i in range(num_quarters):
                qtr = p_qtr if p_qtr else i + 1
                start_month = (qtr * 3) - 2
                end_date = current_quarter_end(p_year, start_month)

                data_quarter = gnc_session.get_account_assets(ASSET_ACCTS, end_date)
                data_quarter[QTR] = str(qtr)

                self.gnucash_data.append(data_quarter)

            # no save needed, we're just reading...
            gnc_session.end_session(False)

            if save_gnc:
                fname = F"updateAssets_gnc-data-{p_year}{('-Q' + str(p_qtr) if p_qtr else '')}"
                save_to_json(fname, self.gnucash_data)

        except Exception as gnce:
            lgr.error(F"Exception: {repr(gnce)}!")
            if gnc_session:
                gnc_session.check_end_session(locals())
            raise gnce

    def fill_google_cell(self, p_col:str, p_row:int, p_time:str):
        self.ggl_update.fill_cell(self.dest, p_col, p_row, p_time)

    def fill_google_data(self, p_year:int, save_ggl:bool):
        """
        >> NOT really necessary to have a separate variable for the Gnucash data, but useful to have all
           the Gnucash data in a separate dict instead of just preparing a Google data dict
        Fill the data list.
        for each item in results, either 1 for one quarter or 4 for four quarters:
            create 5 cell_data's, one each for REV, BAL, CONT, NEC, DEDNS:
            fill in the range based on the year and quarter
            range = SHEET_NAME + '!' + calculated cell
            fill in the values based on the sheet being updated and the type of cell_data
            REV string is '= ${INV} + ${OTH} + ${SAL}'
            DEDNS string is '= ${Mk-Dedns} + ${Lu-Dedns} + ${ML-Dedns}'
            others are just the string from the item
        :param   p_year: year to update
        :param save_ggl: save Google data to a JSON file
        """
        lgr.info(F"UpdateAssets.fill_google_data({p_year},{save_ggl})\n")

        try:
            year_row = BASE_ROW + year_span(p_year, BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
            # get exact row from Quarter value in each item
            for item in self.gnucash_data:
                lgr.info(F"{QTR} = {item[QTR]}")
                int_qtr = int(item[QTR])
                dest_row = year_row + ((int_qtr - 1) * QTR_SPAN)
                lgr.info(F"dest_row = {dest_row}\n")
                for key in item:
                    if key != QTR:
                        # FOR YEAR 2015 OR EARLIER: GET RESP INSTEAD OF Rewards for COLUMN O
                        if key == RESP and p_year > 2015:
                            continue
                        if key == RWRDS and p_year < 2016:
                            continue
                        self.fill_google_cell(ASSET_COLS[key], dest_row, item[key])

            # fill date & time of this update
            today_row = BASE_ROW + 1 + year_span(now_dt.year + 2, BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
            self.fill_google_cell(ASSET_COLS[DATE], today_row, now_dt.strftime(CELL_DATE_STR))
            self.fill_google_cell(ASSET_COLS[DATE], today_row + 1, now_dt.strftime(CELL_TIME_STR))

            str_qtr = None
            if len(self.gnucash_data) == 1:
                str_qtr = self.gnucash_data[0][QTR]

            if save_ggl:
                fname = F"updateAssets_google-data-{str(p_year)}{('-Q' + str_qtr if str_qtr else '')}"
                save_to_json(fname, self.get_google_data())

        except Exception as fgde:
            lgr.error(F"Exception: {repr(fgde)}!")
            raise fgde

# END class UpdateAssets


def process_args() -> ArgumentParser:
    arg_parser = ArgumentParser(description='Update the Assets section of my Google Sheet', prog='updateAssets.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST,SEND+'1',SEND+'2'],
                          help='SEND to Google sheet (1 or 2) OR just TEST')
    required.add_argument('-y', '--year', required=True, help=F"year to update: {BASE_YEAR}..2019")
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices=['1','2','3','4'], help="quarter to update: 1..4")
    arg_parser.add_argument('--gnc_save',  action='store_true', help='Write the Gnucash formatted data to a JSON file')
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google formatted data to a JSON file')
    arg_parser.add_argument('--debug', action='store_true', help='GENERATE DEBUG OUTPUT: MANY LINES!')

    return arg_parser


def process_input_parameters(argl:list) -> (str, bool, bool, bool, str, int, int) :
    args = process_args().parse_args(argl)
    lgr.info(F"\nargs = {args}")

    if args.debug:
        lgr.warning('Printing ALL Debug output!!')

    if not osp.isfile(args.gnucash_file):
        msg = F"File path '{args.gnucash_file}' DOES NOT exist! Exiting..."
        lgr.warning(msg)
        raise Exception(msg)

    lgr.info(F"\nGnucash file = {args.gnucash_file}")

    year = get_int_year(args.year, BASE_YEAR)
    qtr = 0 if args.quarter is None else get_int_quarter(args.quarter)

    return args.gnucash_file, args.gnc_save, args.ggl_save, args.debug, args.mode, year, qtr


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_assets_main(args:list) -> dict :
    lgr.info(F"Parameters = \n{json.dumps(args, indent=4)}")
    gnucash_file, save_gnc, save_ggl, debug, mode, target_year, target_qtr = process_input_parameters(args)

    ua_now = dt.now().strftime(FILE_DATE_FORMAT)
    lgr.info(F"update_balance_main(): Runtime = {ua_now}")

    try:
        updater = UpdateAssets(gnucash_file, mode, debug)

        # either for One Quarter or for Four Quarters if updating an entire Year
        updater.prepare_gnucash_data(save_gnc, target_year, target_qtr)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_data(target_year, save_ggl)

        # send data if in PROD mode
        if SEND in mode:
            response = updater.ggl_update.send_sheets_data()
            fname = F"updateAssets_response-{target_year}{('-Q' + str(target_qtr) if target_qtr else '')}"
            save_to_json(fname, response, ua_now)
        else:
            response = saved_log_info

    except Exception as ae:
        msg = repr(ae)
        lgr.warning(msg)
        response = {F"update_assets_main() EXCEPTION: {msg}"}

    lgr.info(" >>> PROGRAM ENDED.\n")
    if not MULTI_LOGGING:
        finish_logging(GOOGLE_BASENAME, ua_now)
    return response


if __name__ == "__main__":
    MULTI_LOGGING = False
    print('MULTI_LOGGING = False')
    update_assets_main(argv[1:])
