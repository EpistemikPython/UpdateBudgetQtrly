#
# updateBQ_year.py -- use the Gnucash and Google APIs to update my BudgetQtrly document for a specified year
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-03-30
# @updated 2019-03-30

from sys import argv, stdout
from datetime import date, timedelta, datetime
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
from gnucash import Session, GncNumeric

# find the proper path to the account in the gnucash file
REV_ACCTS = {
    "Invest" : ["REV_Invest"],
    "Other"  : ["REV_Other"],
    "Salary" : ["REV_Salary"]
}

# a dictionary with a period name as key, and number of months in that kind of period as the value
PERIODS = {
    "quarterly":  3
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


# noinspection PyUnresolvedReferences
def update_bq_year_main():
    global gnucash_session
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print("NOT ENOUGH parameters!")
        print("usage: {} <book url> <year> <quarter>".format(exe))
        print("PROGRAM EXIT!")
        return

    print("\nrunning {} at run-time: {}\n".format(exe, str(datetime.now())))

    try:
        (gnucash_file, str_year, str_quarter) = argv[1:4]
        print("find Total Revenue in {} for {}-Q{}".format(gnucash_file, str_year, str_quarter))

        rev_year, rev_quarter = [int(blah) for blah in (str_year, str_quarter)]

        start_month = (rev_quarter * 3) - 2

        # a list of all the periods of interest
        # for each period keep the start date, end date, debits and credits sums and overall total
        period_list = [
            [
                start_date, end_date,
                ZERO,  # debits sum
                ZERO,  # credits sum
                ZERO   # TOTAL
            ]
            for start_date, end_date in generate_period_boundaries(rev_year, start_month, 'quarterly', 1)
        ]
        # print(period_list)
        # a copy of the above list with just the period start dates
        period_starts = [e[0] for e in period_list]
        # print(period_starts)

        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()

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
            print("{} Revenue for {}-Q{} = ${}".format(acct_name, str_year, str_quarter, sum_revenue))

        tot_revenue = period_list[0][4] * (-1)
        print("\n{} Revenue for {}-Q{} = ${}".format("TOTAL", str_year, str_quarter, tot_revenue))

        # no save needed, we're just reading..
        gnucash_session.end()
    except Exception as qe:
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()
        raise

    print("\n >>> PROGRAM ENDED.")


if __name__ == "__main__":
    update_bq_year_main()
