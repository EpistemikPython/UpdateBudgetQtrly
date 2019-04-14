#
# updateRevExps.py -- use the Gnucash and Google APIs to update the Assets
#                     in my BudgetQtrly document for a specified year or quarter
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets example
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-06
# @updated 2019-04-07

from sys import argv
from datetime import datetime as dt
from gnucash import Session
import copy
from googleapiclient.discovery import build
from updateCommon import *

# constant strings
AU    = 'Gold'
AG    = 'Silver'
CASH  = 'Cash'
BANK  = 'Bank'
RWRDS = 'Rewards'
RESP  = 'RESP'
OPEN  = 'OPEN'
RRSP  = 'RRSP'
TFSA  = 'TFSA'
HOUSE = 'House'

# find the proper path to the accounts in the gnucash file
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

ASSET_COLS = {
    AU    : 'U',
    AG    : 'T',
    CASH  : 'R',
    BANK  : 'Q',
    RWRDS : 'O',
    RESP  : 'O',
    OPEN  : 'L',
    RRSP  : 'M',
    TFSA  : 'N',
    HOUSE : 'I'
}
BASE_YEAR = 2007
# number of rows between quarters in the same year
QTR_SPAN = 1
# number of rows between years
BASE_YEAR_SPAN = 4

# either for One Quarter or for Four Quarters if updating an entire Year
gnc_data = list()
google_data = list()

now = dt.now().strftime("%Y-%m-%dT%H-%M-%S")


def year_span(year):
    """
    For Asset rows, have to factor in the header row placed every three years
    :param year: int: year to calculate for
    :return: int: value to use to calculate which row to update
    """
    year_diff = year - BASE_YEAR
    return (year_diff * BASE_YEAR_SPAN) + (year_diff // 3)


def get_acct_bal(acct, idate, cur):
    """
    get the balance in the account on this date in this currency
    :param acct: Gnucash account
    :param idate: Date
    :param cur: Gnucash commodity
    :return: python Decimal with balance
    """
    # CALLS ARE RETRIEVING ASSET VALUES FROM DAY BEFORE!!??
    idate += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(idate)
    acct_comm = acct.GetCommodity()
    # print_info("acct_comm = {}".format(acct_comm))
    if acct_comm == cur:
        acct_cad = acct_bal
    else:
        acct_cad = acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, cur, idate)
    # print_info("{} balance on {} = {}".format(acct.GetName(), idate, acct_cad))

    return gnc_numeric_to_python_decimal(acct_cad)


def get_acct_totals(root_account, end_date, cur):
    """
    Get REVENUE data for the specified account for the specified quarter
    :param root_account: Gnucash account: from the Gnucash book
    :param end_date: date: read the account total at the end of the quarter
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: string with sum of totals
    """
    data = {}
    for item in ASSET_ACCTS:
        path = ASSET_ACCTS[item]
        acct = account_from_path(root_account, path)
        acct_name = acct.GetName()

        # get the split amounts for the parent account
        acct_sum = get_acct_bal(acct, end_date, cur)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(acct_name))
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += get_acct_bal(sub_acct, end_date, cur)

        str_sum = acct_sum.to_eng_string()
        print_info("Assets for {} on {} = ${}\n".format(acct_name, end_date, str_sum), MAGENTA)
        data[item] = str_sum

    return data


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_assets(gnucash_file, re_year, re_quarter):
    """
    Get ASSET data for ONE specified Quarter or ALL four Quarters for the specified Year
    :param gnucash_file: string: name of file used to read the values
    :param re_year: int: year to update
    :param re_quarter: int: 1..4 for quarter to update or 0 if updating entire year
    :return: nil
    """
    num_quarters = 1 if re_quarter else 4
    print_info("find Assets in {} for {}{}".format(gnucash_file, re_year, ('-Q' + str(re_quarter)) if re_quarter else ''), GREEN)

    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        commod_tab = gnucash_session.book.get_table()
        # noinspection PyPep8Naming
        CAD = commod_tab.lookup("ISO4217", "CAD")

        for i in range(num_quarters):
            qtr = re_quarter if re_quarter else i + 1

            start_month = (qtr * 3) - 2
            end_date = period_end(re_year, start_month)

            data_quarter = get_acct_totals(root_account, end_date, CAD)
            data_quarter[QTR] = str(qtr)

            gnc_data.append(data_quarter)

        # no save needed, we're just reading...
        gnucash_session.end()

        save_to_json('out/updateAssets_gnc-data', now, gnc_data)

    except Exception as ge:
        print_error("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(177)


def fill_assets_data(mode, re_year):
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
    :param mode: string: 'xxx[prod][send]'
    :param re_year: int: year to update
    :return: data list
    """
    # FOR YEAR 2015 OR EARLIER: GET RESP INSTEAD OF Rewards for COLUMN O
    print_info("\nfill_assets_data({}, {})\n".format(mode, re_year), CYAN)

    dest = QTR_ASTS_2_SHEET
    if '1' in mode:
        dest = QTR_ASTS_SHEET
    print_info("dest = {}\n".format(dest))

    year_row = BASE_ROW + year_span(re_year)
    # get exact row from Quarter value in each item
    for item in gnc_data:
        print_info("{} = {}".format(QTR, item[QTR]))
        int_qtr = int(item[QTR])
        dest_row = year_row + ((int_qtr - 1) * QTR_SPAN)
        print_info("dest_row = {}\n".format(dest_row))
        for key in item:
            if key != QTR:
                if key == RESP and re_year > 2015:
                    continue
                if key == RWRDS and re_year < 2016:
                    continue
                cell = {}
                col = ASSET_COLS[key]
                val = item[key]
                cell_locn = dest + '!' + col + str(dest_row)
                cell['range']  = cell_locn
                cell['values'] = [[val]]
                print_info("cell = {}".format(cell))
                google_data.append(cell)

    save_to_json('out/updateAssets_google-data', now, google_data)
    return google_data


def send_assets(mode, re_year):
    """
    Fill the data list and send to the document
    :param mode: string: 'xxx[prod][send]'
    :param re_year: int: year to update
    :return: server response
    """
    print_info("\nsend_assets({}, {})".format(mode, re_year))

    response = 'NO SEND'
    try:
        fill_assets_data(mode, re_year)

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
            save_to_json('out/updateAssets_response', now, response)

    except Exception as se:
        print_error("Exception: {}!".format(se))
        exit(275)

    return response


def update_assets_main():
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: nil
    """
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print_error("NOT ENOUGH parameters!")
        print_info("usage: {} <book url> mode=<.?[send]1|2> <year> [quarter]".format(exe), GREEN)
        print_error("PROGRAM EXIT!")
        return

    gnucash_file = argv[1]
    mode = argv[2].lower()
    print_info("\nrunning '{}' on '{}' in mode '{}' at run-time: {}\n".format(exe, gnucash_file, mode, now), GREEN)

    re_year = int(argv[3])
    re_quarter = int(argv[4]) if len(argv) > 4 else 0

    get_assets(gnucash_file, re_year, re_quarter)

    send_assets(mode, re_year)

    print_info("\n >>> PROGRAM ENDED.", MAGENTA)


if __name__ == "__main__":
    update_assets_main()
