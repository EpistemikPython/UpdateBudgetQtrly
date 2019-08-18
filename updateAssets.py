##############################################################################################################################
# coding=utf-8
#
# updateAssets.py -- use the Gnucash and Google APIs to update the Assets
#                    in my BudgetQtrly document for a specified year or quarter
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-06'
__updated__ = '2019-08-12'

from gnucash import Session, gnucash_core
from googleapiclient.discovery import build
from updateCommon import *

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

BASE_YEAR:int = 2007
# number of rows between same quarter in adjacent years
BASE_YEAR_SPAN:int = 4
# number of rows between quarters in the same year
QTR_SPAN:int = 1
# number of year groups between header rows
HDR_SPAN:int = 3


# noinspection PyUnresolvedReferences
def get_acct_bal(acct:gnucash_core.Account, p_date:date, p_currency:gnucash_core.GncCommodity):
    """
    get the balance in this account on this date in this currency
    :param       acct: to find
    :param     p_date: to use
    :param p_currency: to use
    :return: python Decimal with balance
    """
    # CALLS ARE RETRIEVING ASSET VALUES FROM DAY BEFORE!!??
    p_date += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(p_date)
    acct_comm = acct.GetCommodity()
    # print_info("acct_comm = {}".format(acct_comm))
    if acct_comm == p_currency:
        acct_cad = acct_bal
    else:
        acct_cad = acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, p_currency, p_date)
    # print_info("{} balance on {} = {}".format(acct.GetName(), idate, acct_cad))

    return gnc_numeric_to_python_decimal(acct_cad)


def get_acct_assets(root_account:gnucash_core.Account, end_date:date, p_currency:gnucash_core.GncCommodity):
    """
    Get ASSET data for the specified account for the specified quarter
    :param root_account: Gnucash Account from the Gnucash book
    :param     end_date: read the account total at the end of the quarter
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :return: string with sum of totals
    """
    data = {}
    for item in ASSET_ACCTS:
        path = ASSET_ACCTS[item]
        acct = account_from_path(root_account, path)
        acct_name = acct.GetName()

        # get the split amounts for the parent account
        acct_sum = get_acct_bal(acct, end_date, p_currency)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print_info("Descendants of {}:".format(acct_name))
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += get_acct_bal(sub_acct, end_date, p_currency)

        str_sum = acct_sum.to_eng_string()
        print_info("Assets for {} on {} = ${}\n".format(acct_name, end_date, str_sum), MAGENTA)
        data[item] = str_sum

    return data


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_gnucash_data(gnucash_file:str, p_year:int, p_qtr:int):
    """
    Get ASSET data for ONE specified Quarter or ALL four Quarters for the specified Year
    :param gnucash_file: name of file used to read the values
    :param       p_year: year to update
    :param        p_qtr: 1..4 for quarter to update or 0 if updating ALL FOUR quarters
    :return: list of Gnucash data
    """
    print_info("find Assets in {} for {}{}".format(gnucash_file, p_year, ('-Q' + str(p_qtr)) if p_qtr else ''), GREEN)
    num_quarters = 1 if p_qtr else 4
    gnc_data = list()
    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        commod_tab = gnucash_session.book.get_table()
        # noinspection PyPep8Naming
        CAD = commod_tab.lookup("ISO4217", "CAD")

        for i in range(num_quarters):
            qtr = p_qtr if p_qtr else i + 1
            start_month = (qtr * 3) - 2
            end_date = current_quarter_end(p_year, start_month)

            data_quarter = get_acct_assets(root_account, end_date, CAD)
            data_quarter[QTR] = str(qtr)

            gnc_data.append(data_quarter)

        # no save needed, we're just reading...
        gnucash_session.end()

        fname = "out/updateAssets_gnc-data-{}{}".format(p_year, ('-Q' + str(p_qtr)) if p_qtr else '')
        save_to_json(fname, now, gnc_data)
        return gnc_data

    except Exception as ge:
        print_error("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(155)


def fill_google_data(mode:str, p_year:int, gnc_data:list):
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
    :param     mode: '.?[send][1]'
    :param   p_year: year to update
    :param gnc_data: Gnucash data for each needed quarter
    :return: data list
    """
    print_info("fill_google_data({}, {}, gnc_data)\n".format(mode, p_year))
    dest = QTR_ASTS_2_SHEET
    if '1' in mode:
        dest = QTR_ASTS_SHEET

    google_data = list()
    try:
        year_row = BASE_ROW + year_span(p_year, BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
        # get exact row from Quarter value in each item
        for item in gnc_data:
            print_info("{} = {}".format(QTR, item[QTR]))
            int_qtr = int(item[QTR])
            dest_row = year_row + ((int_qtr - 1) * QTR_SPAN)
            print_info("dest_row = {}\n".format(dest_row))
            for key in item:
                if key != QTR:
                    # FOR YEAR 2015 OR EARLIER: GET RESP INSTEAD OF Rewards for COLUMN O
                    if key == RESP and p_year > 2015:
                        continue
                    if key == RWRDS and p_year < 2016:
                        continue
                    fill_cell(dest, ASSET_COLS[key], dest_row, item[key], google_data)

        # fill update date & time
        today_row = BASE_ROW + 1 + year_span(today.year+2, BASE_YEAR, BASE_YEAR_SPAN, HDR_SPAN)
        fill_cell(dest, ASSET_COLS[DATE], today_row, today.strftime(FILE_DATE_STR), google_data)
        fill_cell(dest, ASSET_COLS[DATE], today_row+1, today.strftime(CELL_TIME_STR), google_data)

        str_qtr = None
        if len(gnc_data) == 1:
            str_qtr = gnc_data[0][QTR]
        fname = "out/updateAssets_google-data-{}{}".format(str(p_year), ('-Q' + str_qtr) if str_qtr else '')
        save_to_json(fname, now, google_data)

    except Exception as fgde:
        print_error("Exception: {}!".format(repr(fgde)))
        exit(210)

    return google_data


def send_google_data(mode:str, google_data:list):
    """
    Fill the data list and send to the document
    :param        mode: '.?[send][1]'
    :param google_data: Gnucash data for each needed quarter
    :return: server response
    """
    print_info("send_google_data({}, google_data)\n".format(mode))
    response = None
    try:
        assets_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': google_data
        }

        if 'send' in mode:
            creds = get_credentials()
            service = build('sheets', 'v4', credentials=creds)
            vals = service.spreadsheets().values()
            response = vals.batchUpdate(spreadsheetId=get_budget_id(), body=assets_body).execute()

            print_info('\n{} cells updated!'.format(response.get('totalUpdatedCells')))

    except Exception as sgde:
        print_error("Exception: {}!".format(repr(sgde)))
        exit(240)

    return response


def update_assets_main(args:list):
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: string
    """
    if len(args) < 3:
        print_error("NOT ENOUGH parameters!")
        usage = 'usage: py36 updateAssets.py <book url> mode=<.?[send][1]> <year> [quarter]'
        print_info(usage, GREEN)
        print_error("PROGRAM EXIT!")
        return usage

    gnucash_file = args[0]
    mode = args[1].lower()
    print_info("\nrunning in mode '{}' at run-time: {}\n".format(mode, now), CYAN)

    target_year = get_int_year(args[2], BASE_YEAR)
    target_qtr  = get_int_quarter(args[3]) if len(args) > 3 else 0

    # either for One Quarter or for Four Quarters if updating an entire Year
    gnc_data = get_gnucash_data(gnucash_file, target_year, target_qtr)

    google_data = fill_google_data(mode, target_year, gnc_data)
    response = send_google_data(mode, google_data)

    print_info("\n >>> PROGRAM ENDED.", MAGENTA)

    if response:
        fname = "out/updateAssets_response-{}{}".format(target_year, ('-Q' + str(target_qtr)) if target_qtr else '')
        save_to_json(fname, now, response)
        return response
    else:
        return gnc_data


if __name__ == "__main__":
    import sys
    update_assets_main(sys.argv[1:])
