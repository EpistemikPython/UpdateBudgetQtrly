##############################################################################################################################
# coding=utf-8
#
# gnucash_utilities.py -- useful classes, functions & constants
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-07'
__updated__ = '2019-09-07'

from sys import stdout, path
path.append("/home/marksa/dev/git/Python/Utilities/")
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
from gnucash import Session, GncNumeric, Account, GncCommodity, GncCommodityTable
from python_utilities import *


ZERO: Decimal = Decimal(0)
gnc_color = MAGENTA


def begin_session(p_filename:str, p_new:bool) -> (Session, Account, GncCommodityTable):
    gnucash_session = Session(p_filename, is_new=p_new)
    root_account = gnucash_session.book.get_root_account()
    commod_table = gnucash_session.book.get_table()
    return gnucash_session, root_account, commod_table


def end_session(p_session:Session, p_save:bool):
    if p_save:
        p_session.save()
    p_session.end()
    # not needed?
    # p_session.destroy()


def check_end_session(p_session:Session, p_locals:dict):
    if "gnucash_session" in p_locals and p_session is not None:
        p_session.end()


def gnc_numeric_to_python_decimal(numeric:GncNumeric, p_debug:bool=False) -> Decimal:
    """
    convert a GncNumeric value to a python Decimal value
    :param numeric: value to convert
    :param p_debug: debug printing
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.gnc_numeric_to_python_decimal()", gnc_color)

    negative = numeric.negative_p()
    sign = 1 if negative else 0

    val = GncNumeric(numeric.num(), numeric.denom())
    result = val.to_decimal(None)
    if not result:
        raise Exception("GncNumeric value '{}' CANNOT be converted to decimal!".format(val.to_string()))

    digit_tuple = tuple(int(char) for char in str(val.num()) if char != '-')
    denominator = val.denom()
    exponent = int(log10(denominator))

    assert( (10 ** exponent) == denominator )
    return Decimal((sign, digit_tuple, -exponent))


def get_account_balance(acct:Account, p_date:date, p_currency:GncCommodity, p_debug:bool=False) -> Decimal:
    """
    get the BALANCE in this account on this date in this currency
    :param       acct: Gnucash Account
    :param     p_date: required
    :param p_currency: Gnucash commodity
    :param    p_debug: debug printing
    :return: Decimal with balance
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.get_account_balance()", gnc_color)

    # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
    p_date += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(p_date)
    acct_comm = acct.GetCommodity()
    # check if account is already in the desired currency and convert if necessary
    acct_cur = acct_bal if acct_comm == p_currency \
                        else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, p_currency, p_date)

    return gnc_numeric_to_python_decimal(acct_cur)


def get_total_balance(p_root:Account, p_path:list, p_date:date, p_currency:GncCommodity, p_debug:bool=False) -> (str, Decimal):
    """
    get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
    :param  p_root: Gnucash Account from the Gnucash book
    :param     p_path: path to the account
    :param     p_date: to get the balance
    :param p_currency: Gnucash Commodity: currency to use for the totals
    :param    p_debug: debug printing
    :return: string, int: account name and account sum
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.get_total_balance()", gnc_color)

    acct = account_from_path(p_root, p_path)
    acct_name = acct.GetName()
    # get the split amounts for the parent account
    acct_sum = get_account_balance(acct, p_date, p_currency)
    descendants = acct.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        for sub_acct in descendants:
            # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
            acct_sum += get_account_balance(sub_acct, p_date, p_currency)

    if p_debug: SattoLog.print_text("gnucash_utilities.get_total_balance(): {} on {} = ${}"
                                    .format(acct_name, p_date, acct_sum), gnc_color)
    return acct_name, acct_sum


def get_account_assets(p_root:Account, asset_accts:dict, end_date:date, p_currency:GncCommodity, p_debug:bool=False) -> dict:
    """
    Get ASSET data for the specified account for the specified quarter
    :param  asset_accts:
    :param p_root: Gnucash Account from the Gnucash book
    :param     end_date: read the account total at the end of the quarter
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :param      p_debug: debug printing
    :return: string with sum of totals
    """
    data = {}
    for item in asset_accts:
        acct_path = asset_accts[item]
        acct = account_from_path(p_root, acct_path)
        acct_name = acct.GetName()

        # get the split amounts for the parent account
        acct_sum = get_account_balance(acct, end_date, p_currency)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print_info("Descendants of {}:".format(acct_name))
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += get_account_balance(sub_acct, end_date, p_currency)

        str_sum = acct_sum.to_eng_string()
        if p_debug: SattoLog.print_text("Assets for {} on {} = ${}\n".format(acct_name, end_date, str_sum), MAGENTA)
        data[item] = str_sum

    return data


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def account_from_path(top_account:Account, account_path:list, original_path:list=None, p_debug:bool=False) -> Account:
    """
    RECURSIVE function to get a Gnucash Account: starting from the top account and following the path
    :param   top_account: base Account
    :param  account_path: path to follow
    :param original_path: original call path
    :param       p_debug: debug printing
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.account_from_path({}:{})"
                                    .format(top_account.GetName(), account_path), gnc_color)

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


def get_splits(acct:Account, period_starts:list, periods:list, p_debug:bool=False):
    """
    get the splits for the account and each sub-account and add to periods
    :param          acct: to get splits
    :param period_starts: start date for each period
    :param       periods: fill with splits for each quarter
    :param       p_debug: debug printing
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.get_splits()", gnc_color)

    # insert and add all splits in the periods of interest
    for split in acct.GetSplitList():
        trans = split.parent
        # GetDate() returns a datetime but need a date
        trans_date = trans.GetDate().date()

        # use binary search to find the period that starts before or on the transaction date
        period_index = bisect_right(period_starts, trans_date) - 1

        # ignore transactions with a date before the matching period start and after the last period_end
        if period_index >= 0 and trans_date <= periods[len(periods) - 1][1]:
            # get the period bucket appropriate for the split in question
            period = periods[period_index]
            assert( period[1] >= trans_date >= period[0] )

            split_amount = gnc_numeric_to_python_decimal(split.GetAmount())

            # if the amount is negative this is a credit, else a debit
            debit_credit_offset = 1 if split_amount < ZERO else 0

            # add the debit or credit to the sum, using the offset to get in the right bucket
            period[2 + debit_credit_offset] += split_amount

            # add the debit or credit to the overall total
            period[4] += split_amount


def fill_splits(p_root:Account, target_path:list, period_starts:list, periods:list, p_debug:bool=False) -> str :
    """
    fill the period list for each account
    :param     p_root: from the Gnucash book
    :param   target_path: account hierarchy from root account to target account
    :param period_starts: start date for each period
    :param       periods: fill with the splits dates and amounts for requested time span
    :param       p_debug: debug printing
    :return: name of target_acct
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.fill_splits()", gnc_color)

    account_of_interest = account_from_path(p_root, target_path)
    acct_name = account_of_interest.GetName()
    if p_debug: SattoLog.print_text("\naccount_of_interest = {}".format(acct_name), gnc_color)

    # get the split amounts for the parent account
    get_splits(account_of_interest, period_starts, periods)
    descendants = account_of_interest.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        # print("Descendants of {}:".format(account_of_interest.GetName()))
        for subAcct in descendants:
            # print("{} balance = {}".format(subAcct.GetName(), gnc_numeric_to_python_decimal(subAcct.GetBalance())))
            get_splits(subAcct, period_starts, periods)

    csv_write_period_list(periods)
    return acct_name


def csv_write_period_list(periods:list, p_debug:bool=False):
    """
    Write out the details of the submitted period list in csv format
    :param periods: dates and amounts for each quarter
    :param p_debug: debug printing
    :return: to stdout
    """
    if p_debug: SattoLog.print_text("gnucash_utilities.csv_write_period_list()", gnc_color)

    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in periods:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))


def show_account(p_root:Account, p_path:list, logger:SattoLog, p_color:str):
    """
    display an account and its descendants
    :param  p_root: Gnucash root
    :param  p_path: to the account
    :param  logger: to use for printing
    :param p_color: to print with
    :return: nil
    """
    acct = account_from_path(p_root, p_path)
    acct_name = acct.GetName()
    logger.print_info("account = {}".format(acct_name), p_color)
    descendants = acct.get_descendants()
    if len(descendants) == 0 :
        logger.print_info("{} has NO Descendants!".format(acct_name), p_color)
    else :
        logger.print_info("Descendants of {}:".format(acct_name), p_color)
        # for subAcct in descendants:
        # print_info("{}".format(subAcct.GetName()))
