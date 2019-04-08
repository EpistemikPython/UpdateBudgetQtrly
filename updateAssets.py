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
    OPEN  : ["FAMILY", "INVEST", OPEN],
    RRSP  : ["FAMILY", "INVEST", RRSP],
    TFSA  : ["FAMILY", "INVEST", TFSA],
    HOUSE : ["FAMILY", HOUSE]
}

# store the values needed to update the document
ASSET_RESULTS = {
    QTR   : '0',
    AU    : '0',
    AG    : '0',
    CASH  : '0',
    BANK  : '0',
    RWRDS : '0',
    OPEN  : '0',
    RRSP  : '0',
    TFSA  : '0',
    HOUSE : '0'
}

ASSET_COLS = {
    AU    : 'U',
    AG    : 'T',
    CASH  : 'R',
    BANK  : 'Q',
    RWRDS : 'O',
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
    str_sum = 'EMPTY'
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
        ASSET_RESULTS[item] = str_sum

    return str_sum


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
            acct_total = ZERO
            qtr = re_quarter if re_quarter else i + 1
            ASSET_RESULTS[QTR] = str(qtr)

            start_month = (qtr * 3) - 2
            end_date = period_end(re_year, start_month)

            acct_total = get_acct_totals(root_account, end_date, CAD)
            # print("\n{} Assets for {}-Q{} = ${}".format("TOTAL", re_year, qtr, acct_total))

            results.append(copy.deepcopy(ASSET_RESULTS))
            # print(json.dumps(ASSET_RESULTS, indent=4))

        # no save needed, we're just reading...
        gnucash_session.end()

        save_to_json('out/updateAssets_results', now, results)

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
    print_info("\nfill_assets_data({}, {})\n".format(mode, re_year), CYAN)

    dest = QTR_ASTS_2_SHEET
    if 'prod' in mode:
        dest = QTR_ASTS_SHEET
    print_info("dest = {}\n".format(dest))

    year_row = BASE_ROW + year_span(re_year)
    # get exact row from Quarter value in each item
    for item in results:
        print_info("{} = {}".format(QTR, item[QTR]))
        int_qtr = int(item[QTR])
        dest_row = year_row + ((int_qtr - 1) * QTR_SPAN)
        print_info("dest_row = {}\n".format(dest_row))
        for key in item:
            if key != QTR:
                cell = copy.copy(cell_data)
                col = ASSET_COLS[key]
                val = item[key]
                cell_locn = dest + '!' + col + str(dest_row)
                cell['range']  = cell_locn
                cell['values'] = [[val]]
                print_info("cell = {}".format(cell))
                data.append(cell)
    return data


def send_assets(mode, re_year):
    """
    Fill the data list and send to the document
    :param mode: string: 'xxx[prod][send]'
    :param re_year: int: year to update
    :return: server response
    """
    print_info("\nsend_assets({}, {})".format(mode, re_year))
    print_info("cell_data['range'] = {}".format(cell_data['range']))
    print_info("cell_data['values'][0][0] = {}".format(cell_data['values'][0][0]))

    response = 'NO SEND'
    try:
        fill_assets_data(mode, re_year)

        save_to_json('out/updateAssets_data', now, data)

        creds = get_credentials()

        service = build('sheets', 'v4', credentials=creds)
        srv_sheets = service.spreadsheets()

        my_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }

        if 'send' in mode:
            vals = srv_sheets.values()
            response = vals.batchUpdate(spreadsheetId=BUDGET_QTRLY_ID, body=my_body).execute()

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
        print_info("usage: {} <book url> <mode=xxx[prod][send]> <year> [quarter]".format(exe), GREEN)
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
