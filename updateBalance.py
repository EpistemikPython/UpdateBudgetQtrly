#
# updateBalance.py -- use the Gnucash and Google APIs to update the 'Balance' sheet
#                     in my BudgetQtrly document for today or for a specified year
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets example
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-06
# @updated 2019-04-14

from sys import argv
from datetime import datetime as dt
from gnucash import Session
import copy
from googleapiclient.discovery import build
from updateCommon import *

# constant strings
ASTS  = 'FAMILY'
LIAB  = 'LIABS'
TRUST = 'TRUST'
CHAL  = 'XCHALET'
HOUSE = 'House'
TODAY = 'Today'

# path to the accounts in the Gnucash file
BALANCE_ACCTS = {
    ASTS  : [ASTS],
    LIAB  : [LIAB],
    TRUST : [TRUST],
    CHAL  : [CHAL],
    HOUSE : [ASTS, HOUSE]
}

BALANCE_COLS = {
    LIAB  : {'Yrs': 'U', 'Mths': 'L'},
    TODAY : 'C',
    ASTS  : 'K'
}

BASE_YEAR = 2008
# number of rows between years
BASE_YEAR_SPAN = 1

BASE_MTHLY_ROW = 20

TRUST_RANGE  = 'Balance 1!C21'
CHALET_RANGE = 'Balance 1!C22'
HOUSE_RANGE  = 'Balance 1!C26'
ASSETS_RANGE = 'Balance 1!C27'
LIABS_RANGE  = 'Balance 1!C28'

today = dt.now()
now = today.strftime("%Y-%m-%dT%H-%M-%S")


def year_span(year):
    """
    For Balance rows, have to factor in the header row placed every NINE years
    :param year: int: year to calculate for
    :return: int: value to use to calculate which row to update
    """
    year_diff = year - BASE_YEAR
    return (year_diff * BASE_YEAR_SPAN) + (year_diff // 9)


def get_acct_bal(acct, idate, cur):
    """
    get the BALANCE in the account on this date in this currency
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
    Get BALANCE data for the specified account for the specified quarter
    :param root_account: Gnucash account: from the Gnucash book
    :param end_date: date: read the account total at the end of the quarter
    :param cur: Gnucash Commodity: currency to use for the totals
    :return: string with sum of totals
    """
    data = {}
    for item in BALANCE_ACCTS:
        path = BALANCE_ACCTS[item]
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
def get_gnucash_data(gnucash_file, domain):
    """
    Get Balance data for TODAY:
      LIABS, House, FAMILY, XCHALET, TRUST
    OR for the specified year:
      IF CURRENT YEAR: LIABS for ALL completed months; FAMILY for ALL non-3 completed months in year
      IF PREVIOUS YEAR: LIABS for year, for ALL NON-completed months; FAMILY for ALL non-3 NON-completed months in year
      ELSE: LIABS for year
    :param gnucash_file: string: name of file used to read the values
    :param domain: string: what to update
    :return: nil
    """
    print_info("find Balances in {} for {}".format(gnucash_file, domain), GREEN)
    gnc_data = list()
    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()
        commod_tab = gnucash_session.book.get_table()
        # noinspection PyPep8Naming
        CAD = commod_tab.lookup("ISO4217", "CAD")

        if domain == 'today':
            data = fill_today(root_account)
        else:
            year = get_year(domain)

        if year == today.year:
            data = fill_current_year(root_account)
        elif year - today.year == 1:
            data = fill_previous_year(root_account)
        else:
            data = fill_year(year, root_account)

        gnc_data.append(data)

        # no save needed, we're just reading...
        gnucash_session.end()

        fname = "out/updateBalance_gnc-data-{}".format(domain)
        save_to_json(fname, now, gnc_data)
        return gnc_data

    except Exception as ge:
        print_error("Exception: {}!".format(ge))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        exit(167)


def fill_google_data(mode, re_year, gnc_data):
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
    :param gnc_data: list: Gnucash data for each needed quarter
    :return: data list
    """
    print_info("\nfill_assets_data({}, {})\n".format(mode, re_year), CYAN)

    dest = QTR_ASTS_2_SHEET
    if '1' in mode:
        dest = QTR_ASTS_SHEET
    print_info("dest = {}\n".format(dest))

    google_data = list()
    year_row = BASE_ROW + year_span(re_year)
    # get exact row from Quarter value in each item
    for item in gnc_data:
        print_info("{} = {}".format(QTR, item[QTR]))
        int_qtr = int(item[QTR])
        dest_row = year_row + ((int_qtr - 1) * QTR_SPAN)
        print_info("dest_row = {}\n".format(dest_row))
        for key in item:
            if key != QTR:
                cell = {}
                col = BALANCE_COLS[key]
                val = item[key]
                cell_locn = dest + '!' + col + str(dest_row)
                cell['range']  = cell_locn
                cell['values'] = [[val]]
                print_info("cell = {}".format(cell))
                google_data.append(cell)

    str_qtr = None
    if len(gnc_data) == 1:
        str_qtr = gnc_data[0][QTR]
    fname = "out/updateBalance_google-data-{}{}".format(str(re_year), ('-Q' + str_qtr) if str_qtr else '')
    save_to_json(fname, now, google_data)
    return google_data


def send_google_data(mode, re_year, gnc_data):
    """
    Fill the data list and send to the document
    :param mode: string: 'xxx[prod][send]'
    :param re_year: int: year to update
    :param gnc_data: list: Gnucash data for each needed quarter
    :return: server response
    """
    print_info("\nsend_assets({}, {})".format(mode, re_year))

    response = 'NO SEND'
    try:
        google_data = fill_google_data(mode, re_year, gnc_data)
        save_to_json('out/updateAssets_data', now, google_data)

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

    except Exception as se:
        print_error("Exception: {}!".format(se))
        exit(275)

    return response


def update_balance_main():
    """
    Main: check command line and call functions to get the data from Gnucash book and send to Google document
    :return: nil
    """
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print_error("NOT ENOUGH parameters!")
        print_info("usage: {} <book url> mode=<.?[send]> <year|'today'>".format(exe), GREEN)
        print_error("PROGRAM EXIT!")
        return

    gnucash_file = argv[1]
    mode = argv[2].lower()
    print_info("\nrunning '{}' on '{}' in mode '{}' at run-time: {}\n".format(exe, gnucash_file, mode, now), GREEN)

    domain = argv[3].lower()

    gnc_data = get_gnucash_data(gnucash_file, domain)

    response = send_google_data(mode, gnc_data)
    if response:
        fname = "out/updateBalance_response-{}".format(domain)
        save_to_json(fname, now, response)

    print_info("\n >>> PROGRAM ENDED.", MAGENTA)


if __name__ == "__main__":
    update_balance_main()
