#
# updateRevExps.py -- use the Gnucash and Google APIs to update my BudgetQtrly document for a specified year
#                     in the 'All Inc' and 'Nec Inc' sheets
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-03-30
# @updated 2019-03-31

from sys import argv, stdout
from datetime import date, timedelta, datetime
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
import json
from gnucash import Session, GncNumeric

# constant strings
QTR   = 'quarterly'
YR    = 'yearly'
INV   = 'Invest'
OTH   = 'Other'
SAL   = 'Salary'
BAL   = 'Balance'
CONT  = 'Contingent'
NEC   = 'Necessary'
DEDNS = 'Sal_Dedns'

# find the proper path to the accounts in the gnucash file
REV_ACCTS = {
    INV : ["REV_Invest"],
    OTH : ["REV_Other"],
    SAL : ["REV_Salary"]
}
EXP_ACCTS = {
    BAL   : ["EXP_Balance"],
    CONT  : ["EXP_CONTINGENT"],
    NEC   : ["EXP_NECESSARY"],
    DEDNS : ["EXP_Salary"]
}

# store the sub-totals needed to update the document
REV_EXP_RESULTS = {
    INV   : '0',
    OTH   : '0',
    SAL   : '0',
    BAL   : '0',
    CONT  : '0',
    NEC   : '0',
    DEDNS : '0'
}

# a dictionary with a period name as key, and number of months in that kind of period as the value
PERIODS = {
    QTR :  3 ,
    YR  : 12
}

NUM_MONTHS = 12
ONE_DAY = timedelta(days=1)
ZERO = Decimal(0)


# noinspection PyUnresolvedReferences
def gnc_numeric_to_python_decimal(numeric):
    negative = numeric.negative_p()
    sign = 1 if negative else 0

    copy = GncNumeric(numeric.num(), numeric.denom())
    result = copy.to_decimal(None)
    if not result:
        raise Exception("GncNumeric value '{}' CANNOT be converted to decimal!".format(copy.to_string()))

    digit_tuple = tuple(int(char) for char in str(copy.num()) if char != '-')
    denominator = copy.denom()
    exponent = int(log10(denominator))
    assert( (10 ** exponent) == denominator )
    return Decimal((sign, digit_tuple, -exponent))


def next_period_start(start_year, start_month, period_type):
    # add numbers of months for the period length
    end_month = start_month + PERIODS[period_type]

    # use integer division to find out if the new end month is in a different year,
    # what year it is, and what the end month number should be changed to.
    end_year = start_year + ((end_month - 1) // NUM_MONTHS)
    end_month = ((end_month - 1) % NUM_MONTHS) + 1

    return end_year, end_month


def period_end(start_year, start_month, period_type):
    if period_type not in PERIODS:
        raise Exception("'{}' is NOT a valid period >> MUST be one of '{}'!".format(period_type, str(PERIODS.keys())))

    end_year, end_month = next_period_start(start_year, start_month, period_type)

    # last step, the end date is one day back from the start of the next period
    # so we get a period end like 2010-03-31 instead of 2010-04-01
    return date(end_year, end_month, 1) - ONE_DAY


def generate_period_boundaries(start_year, start_month, period_type, periods):
    for i in range(periods):
        yield( date(start_year, start_month, 1), period_end(start_year, start_month, period_type) )
        start_year, start_month = next_period_start(start_year, start_month, period_type)


def account_from_path(top_account, account_path, original_path=None):
    # print("top_account = %s, account_path = %s, original_path = %s" % (top_account, account_path, original_path))
    if original_path is None:
        original_path = account_path
    account, account_path = account_path[0], account_path[1:]
    # print("account = %s, account_path = %s" % (account, account_path))

    account = top_account.lookup_by_name(account)
    # print("account = " + str(account))
    if account is None:
        raise Exception("Path '" + str(original_path) + "' could NOT be found!")
    if len(account_path) > 0:
        return account_from_path(account, account_path, original_path)
    else:
        return account


def get_splits(acct, period_starts, period_list):
    # insert and add all splits in the periods of interest
    for split in acct.GetSplitList():
        trans = split.parent
        # GetDate() returns a datetime
        tx_datetm = trans.GetDate()
        # convert to a date
        trans_date = tx_datetm.date()

        # use binary search to find the period that starts before or on the transaction date
        period_index = bisect_right(period_starts, trans_date) - 1

        # ignore transactions with a date before the matching period start and after the last period_end
        if period_index >= 0 and trans_date <= period_list[len(period_list) - 1][1]:

            # get the period bucket appropriate for the split in question
            period = period_list[period_index]

            assert( period[1] >= trans_date >= period[0] )

            split_amount = gnc_numeric_to_python_decimal(split.GetAmount())

            # if the amount is negative this is a credit, else a debit
            debit_credit_offset = 1 if split_amount < ZERO else 0

            # add the debit or credit to the sum, using the offset to get in the right bucket
            period[2 + debit_credit_offset] += split_amount

            # add the debit or credit to the overall total
            period[4] += split_amount


def get_revenue(root_account, period_starts, period_list, re_year, str_quarter):
    for item in REV_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = REV_ACCTS[item]
        # print("acct = {}".format(acct_base))

        account_of_interest = account_from_path(root_account, acct_base)
        acct_name = account_of_interest.GetName()
        print("\naccount_of_interest = {}".format(acct_name))
        # get the split amounts for the parent account
        get_splits(account_of_interest, period_starts, period_list)

        descendants = account_of_interest.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(account_of_interest.GetName()))
            for subAcct in descendants:
                # print("{} balance = {}".format(subAcct.GetName(), gnc_numeric_to_python_decimal(subAcct.GetBalance())))
                get_splits(subAcct, period_starts, period_list)

        # write out the column headers
        csv_writer = csv.writer(stdout)
        # csv_writer.writerow('')
        csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

        # write out the overall totals for the account of interest
        for start_date, end_date, debit_sum, credit_sum, total in period_list:
            csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))

        sum_revenue = (period_list[0][2] + period_list[0][3]) * (-1)
        REV_EXP_RESULTS[item] = sum_revenue.to_eng_string()
        print("{} Revenue for {}{} = ${}".format(acct_name, re_year, str_quarter, sum_revenue))


def get_expenses(root_account, period_starts, period_list, re_year, str_quarter):
    for item in EXP_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = EXP_ACCTS[item]
        # print("acct = {}".format(acct_base))

        account_of_interest = account_from_path(root_account, acct_base)
        acct_name = account_of_interest.GetName()
        print("\naccount_of_interest = {}".format(acct_name))

        # get the split amounts for the parent account
        get_splits(account_of_interest, period_starts, period_list)

        descendants = account_of_interest.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(account_of_interest.GetName()))
            for subAcct in descendants:
                # print("{} balance = {}".format(subAcct.GetName(), gnc_numeric_to_python_decimal(subAcct.GetBalance())))
                get_splits(subAcct, period_starts, period_list)

        # write out the column headers
        csv_writer = csv.writer(stdout)
        # csv_writer.writerow('')
        csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

        # write out the overall totals for the account of interest
        for start_date, end_date, debit_sum, credit_sum, total in period_list:
            csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))

        sum_expenses = (period_list[0][2] + period_list[0][3])
        REV_EXP_RESULTS[item] = sum_expenses.to_eng_string()
        print("{} Expenses for {}{} = ${}".format(acct_name.split('_')[-1], re_year, str_quarter, sum_expenses))


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def update_rev_exps_main():
    exe = argv[0].split('/')[-1]
    if len(argv) < 3:
        print("NOT ENOUGH parameters!")
        print("usage: {} <book url> <year> [quarter]".format(exe))
        print("PROGRAM EXIT!")
        return

    print("\nrunning {} at run-time: {}\n".format(exe, str(datetime.now())))

    try:
        gnucash_file = argv[1]
        re_year = int(argv[2])
        period_type = QTR if len(argv) > 3 else YR

        re_quarter = int(argv[3]) if len(argv) > 3 else 0
        start_month = (re_quarter * 3) - 2 if len(argv) > 3 else 1
        str_quarter = '-Q' + str(re_quarter) if re_quarter else ''

        print("find Revenue & Expenses in {} for {}{}".format(gnucash_file, re_year, str_quarter))

        # a list of all the periods of interest
        # for each period keep the start date, end date, debits and credits sums and overall total
        period_list = [
            [
                start_date, end_date,
                ZERO,  # debits sum
                ZERO,  # credits sum
                ZERO  # TOTAL
            ]
            for start_date, end_date in generate_period_boundaries(re_year, start_month, period_type, 1)
        ]
        # print(period_list)
        # a copy of the above list with just the period start dates
        period_starts = [e[0] for e in period_list]
        # print(period_starts)

        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()

        get_revenue(root_account, period_starts, period_list, re_year, str_quarter)
        tot_revenue = period_list[0][4] * (-1)
        print("\n{} Revenue for {}{} = ${}".format("TOTAL", re_year, str_quarter, tot_revenue))

        period_list[0][4] = 0
        get_expenses(root_account, period_starts, period_list, re_year, str_quarter)
        tot_expenses = period_list[0][4]
        print("\n{} Expenses for {}{} = ${}".format("TOTAL", re_year, str_quarter, tot_expenses))

        print(json.dumps(REV_EXP_RESULTS, indent=4))
        # no save needed, we're just reading..

        gnucash_session.end()

    except Exception as qe:
        print("Exception: {}!".format(qe))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        raise

    print("\n >>> PROGRAM ENDED.")


if __name__ == "__main__":
    update_rev_exps_main()
