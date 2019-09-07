##############################################################################################################################
# coding=utf-8
#
# updateAssets.py -- use the Gnucash and Google APIs to update the Assets
#                    in my BudgetQtrly document for a specified year or quarter
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-06'
__updated__ = '2019-09-06'

from sys import path
path.append("/home/marksa/dev/git/Python/Gnucash/createGncTxs")
path.append("/home/marksa/dev/git/Python/Google")
from argparse import ArgumentParser
from gnucash_utilities import *
from google_utilities import *
from investment import *

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


class UpdateAssets:
    def __init__(self, p_filename:str, p_mode:str, p_debug:bool):
        self.log = SattoLog(p_debug)
        self.log.print_info('UpdateAssets', GREEN)
        self.color = CYAN

        self.gnucash_file = p_filename
        self.gnucash_data = []
        # cell(s) with location and value to update on Google sheet
        self.google_data = []

        self.mode = p_mode
        # Google sheet to update
        self.dest = QTR_ASTS_2_SHEET
        if '1' in self.mode:
            self.dest = QTR_ASTS_SHEET
        self.log.print_info("dest = {}".format(self.dest), self.color)

    BASE_YEAR:int = 2007
    # number of rows between same quarter in adjacent years
    BASE_YEAR_SPAN:int = 4
    # number of rows between quarters in the same year
    QTR_SPAN:int = 1
    # number of year groups between header rows
    HDR_SPAN:int = 3

    def get_gnucash_data(self) -> list:
        return self.gnucash_data

    def get_google_data(self) -> list:
        return self.google_data

    def get_log(self) -> list :
        return self.log.get_log()

    def fill_gnucash_data(self, save_gnc:bool, p_year:int, p_qtr:int):
        """
        Get ASSET data for ONE specified Quarter or ALL four Quarters for the specified Year
        :param save_gnc: save gnucash data to json file
        :param   p_year: year to update
        :param    p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
        """
        self.log.print_info("UA.fill_gnucash_data(): find Assets in {} for {}{}"
                            .format(self.gnucash_file, p_year, ('-Q' + str(p_qtr)) if p_qtr else ''), GREEN)
        num_quarters = 1 if p_qtr else 4
        gnucash_session = None
        try:
            gnucash_session, root_account, commod_table = begin_session(self.gnucash_file, False)
            currency = commod_table.lookup("ISO4217", "CAD")

            for i in range(num_quarters):
                qtr = p_qtr if p_qtr else i + 1
                start_month = (qtr * 3) - 2
                end_date = current_quarter_end(p_year, start_month)

                data_quarter = get_account_assets(root_account, ASSET_ACCTS, end_date, currency)
                data_quarter[QTR] = str(qtr)

                self.gnucash_data.append(data_quarter)

            # no save needed, we're just reading...
            end_session(gnucash_session, False)

            if save_gnc:
                fname = "out/updateAssets_gnc-data-{}{}".format(p_year, ('-Q' + str(p_qtr)) if p_qtr else '')
                save_to_json(fname, now, self.gnucash_data)

        except Exception as gnce:
            self.log.print_error("Exception: {}!".format(repr(gnce)))
            check_end_session(gnucash_session, locals())
            raise gnce

    def fill_google_cell(self, p_col:str, p_row:int, p_time:str):
        fill_cell(self.dest, p_col, p_row, p_time, self.google_data)

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
        self.log.print_info("UA.fill_google_data({},{})\n".format(p_year, save_ggl))

        try:
            year_row = BASE_ROW + year_span(p_year, self.BASE_YEAR, self.BASE_YEAR_SPAN, self.HDR_SPAN)
            # get exact row from Quarter value in each item
            for item in self.gnucash_data:
                self.log.print_info("{} = {}".format(QTR, item[QTR]))
                int_qtr = int(item[QTR])
                dest_row = year_row + ((int_qtr - 1) * self.QTR_SPAN)
                self.log.print_info("dest_row = {}\n".format(dest_row))
                for key in item:
                    if key != QTR:
                        # FOR YEAR 2015 OR EARLIER: GET RESP INSTEAD OF Rewards for COLUMN O
                        if key == RESP and p_year > 2015:
                            continue
                        if key == RWRDS and p_year < 2016:
                            continue
                        self.fill_google_cell(ASSET_COLS[key], dest_row, item[key])

            # fill date & time of this update
            today_row = BASE_ROW + 1 + year_span(today.year+2, self.BASE_YEAR, self.BASE_YEAR_SPAN, self.HDR_SPAN)
            self.fill_google_cell(ASSET_COLS[DATE], today_row, today.strftime(FILE_DATE_STR))
            self.fill_google_cell(ASSET_COLS[DATE], today_row+1, today.strftime(CELL_TIME_STR))

            str_qtr = None
            if len(self.gnucash_data) == 1:
                str_qtr = self.gnucash_data[0][QTR]

            if save_ggl:
                fname = "out/updateAssets_google-data-{}{}".format(str(p_year), ('-Q' + str_qtr) if str_qtr else '')
                save_to_json(fname, now, self.google_data)

        except Exception as fgde:
            self.log.print_error("Exception: {}!".format(repr(fgde)))
            raise fgde

# END class UpdateAssets


def process_args() -> ArgumentParser:
    arg_parser = ArgumentParser(description='Update the Assets section of my Google Sheet', prog='updateAssets.py')
    # required arguments
    required = arg_parser.add_argument_group('REQUIRED')
    required.add_argument('-g', '--gnucash_file', required=True, help='path & filename of the Gnucash file to use')
    required.add_argument('-m', '--mode', required=True, choices=[TEST, PROD+'1', PROD+'2'],
                          help='WRITE to Google sheet (1 or 2) OR just TEST (NO write)')
    required.add_argument('-y', '--year', required=True, help="year to update: {}..2019".format(UpdateAssets.BASE_YEAR))
    # optional arguments
    arg_parser.add_argument('-q', '--quarter', choices=['1','2','3','4'], help="quarter to update: 1..4")
    arg_parser.add_argument('--gnc_save',  action='store_true', help='Write the Gnucash formatted data to a JSON file')
    arg_parser.add_argument('--ggl_save',  action='store_true', help='Write the Google formatted data to a JSON file')
    arg_parser.add_argument('--debug', action='store_true', help='GENERATE DEBUG OUTPUT: MANY LINES!')

    return arg_parser


def process_input_parameters(argl:list) -> (str, bool, bool, bool, str, int, int) :
    args = process_args().parse_args(argl)
    SattoLog.print_text("\nargs = {}".format(args), BROWN)

    if args.debug:
        SattoLog.print_text('Printing ALL Debug output!!', RED)

    if not osp.isfile(args.gnucash_file):
        SattoLog.print_text("File path '{}' does not exist! Exiting...".format(args.gnucash_file), RED)
        exit(209)
    SattoLog.print_text("\nGnucash file = {}".format(args.gnucash_file), CYAN)

    year = get_int_year(args.year, UpdateAssets.BASE_YEAR)
    qtr = 0 if args.quarter is None else get_int_quarter(args.quarter)

    return args.gnucash_file, args.gnc_save, args.ggl_save, args.debug, args.mode, year, qtr


# TODO: fill in date column for previous month when updating 'today', check to update 'today' or 'tomorrow'
def update_assets_main(args:list) -> dict :
    SattoLog.print_text("Parameters = \n{}".format(json.dumps(args, indent=4)), GREEN)
    gnucash_file, save_gnc, save_ggl, debug, mode, target_year, target_qtr = process_input_parameters(args)

    ub_now = dt.now().strftime(DATE_STR_FORMAT)
    SattoLog.print_text("update_balance_main(): Runtime = {}".format(ub_now), BLUE)

    try:
        updater = UpdateAssets(gnucash_file, mode, debug)

        # either for One Quarter or for Four Quarters if updating an entire Year
        updater.fill_gnucash_data(save_gnc, target_year, target_qtr)

        # get the requested data from Gnucash and package in the update format required by Google sheets
        updater.fill_google_data(target_year, save_ggl)

        # send data if in PROD mode
        if PROD in mode:
            response = send_sheets_data(updater.get_google_data())
            fname = "out/updateAssets_response-{}{}".format(target_year , ('-Q' + str(target_qtr)) if target_qtr else '')
            save_to_json(fname, ub_now, response)
        else:
            response = updater.get_log()

    except Exception as ae:
        msg = repr(ae)
        SattoLog.print_warning(msg)
        response = {"update_assets_main() EXCEPTION:": "{}".format(msg)}

    SattoLog.print_text(" >>> PROGRAM ENDED.\n", GREEN)
    return response


if __name__ == "__main__":
    from sys import argv
    update_assets_main(argv[1:])
