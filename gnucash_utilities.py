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
__updated__ = '2019-08-30'

from sys import stdout, path
path.append("/home/marksa/dev/git/Python/Utilities/")
from datetime import date, timedelta
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
from gnucash import GncNumeric, Account, GncCommodity
from python_utilities import *

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
TOTAL = 'Total'
ASTS  = 'FAMILY'
LIAB  = 'LIABS'
TRUST = 'TRUST'
CHAL  = 'XCHALET'
DATE  = 'Date'
TODAY = 'Today'
QTR   = 'Quarter'
YR    = 'Year'
MTH   = 'Month'
REV   = 'Revenue'
INV   = 'Invest'
OTH   = 'Other'
EMP   = 'Employment'
BAL   = 'Balance'
CONT  = 'Contingent'
NEC   = 'Necessary'
DEDNS = 'Emp_Dedns'
TRADE = 'Trade'
PRICE = 'Price'
BOTH  = 'Both'
TEST  = 'test'
PROD  = 'prod'

# number of months in a Quarter
PERIOD_QTR:int = 3

YEAR_MONTHS:int = 12
ONE_DAY:timedelta = timedelta(days=1)
ZERO:Decimal = Decimal(0)

# sheet names in Budget Quarterly
ALL_INC_SHEET:str    = 'All Inc 1'
ALL_INC_2_SHEET:str  = 'All Inc 2'
NEC_INC_SHEET:str    = 'Nec Inc 1'
NEC_INC_2_SHEET:str  = 'Nec Inc 2'
BAL_1_SHEET:str      = 'Balance 1'
BAL_2_SHEET:str      = 'Balance 2'
QTR_ASTS_SHEET:str   = 'Assets 1'
QTR_ASTS_2_SHEET:str = 'Assets 2'
ML_WORK_SHEET:str    = 'ML Work'
CALCULNS_SHEET:str   = 'Calculations'

# first data row in Budget-qtrly.gsht
BASE_ROW:int = 3


class GnucashUtilities:
    def __init__(self):
        SattoLog.print_text("GnucashUtilities", GREEN)

    my_color = MAGENTA

    @staticmethod
    def gnc_numeric_to_python_decimal(numeric:GncNumeric) -> Decimal :
        """
        convert a GncNumeric value to a python Decimal value
        :param numeric: value to convert
        """
        # SattoLog.print_text("GnucashUtilities.gnc_numeric_to_python_decimal()", GnucashUtilities.my_color)

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

    @staticmethod
    def get_account_balance(acct:Account, p_date:date, p_currency:GncCommodity) -> Decimal :
        """
        get the BALANCE in this account on this date in this currency
        :param       acct: Gnucash Account
        :param     p_date: required
        :param p_currency: Gnucash commodity
        :return: Decimal with balance
        """
        # SattoLog.print_text("GnucashUtilities.get_account_balance()", GnucashUtilities.my_color)

        # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
        p_date += ONE_DAY

        acct_bal = acct.GetBalanceAsOfDate(p_date)
        acct_comm = acct.GetCommodity()
        # check if account is already in the desired currency and convert if necessary
        acct_cur = acct_bal if acct_comm == p_currency \
                            else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, p_currency, p_date)

        return GnucashUtilities.gnc_numeric_to_python_decimal(acct_cur)

    @staticmethod
    def get_total_balance(root_acct, p_path: list, p_date: date, p_currency):
        """
        get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
        :param  root_acct: Gnucash Account from the Gnucash book
        :param     p_path: path to the account
        :param     p_date: to get the balance
        :param p_currency: Gnucash Commodity: currency to use for the totals
        :return: string, int: account name and account sum
        """
        # SattoLog.print_text("GnucashUtilities.get_total_balance()", GnucashUtilities.my_color)

        acct = GnucashUtilities.account_from_path(root_acct, p_path)
        acct_name = acct.GetName()
        # get the split amounts for the parent account
        acct_sum = GnucashUtilities.get_account_balance(acct, p_date, p_currency)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += GnucashUtilities.get_account_balance(sub_acct, p_date, p_currency)

        SattoLog.print_text("GnucashUtilities.get_total_balance(): {} on {} = ${}".format(acct_name, p_date, acct_sum), GnucashUtilities.my_color)
        return acct_name, acct_sum

    @staticmethod
    def account_from_path(top_account:Account, account_path:list, original_path:list=None) -> Account :
        """
        RECURSIVE function to get a Gnucash Account: starting from the top account and following the path
        :param   top_account: base Account
        :param  account_path: path to follow
        :param original_path: original call path
        """
        SattoLog.print_text("GnucashUtilities.account_from_path({}:{})"
                            .format(top_account.GetName(), account_path), GnucashUtilities.my_color)

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
            return GnucashUtilities.account_from_path(account, account_path, original_path)
        else:
            return account

    @staticmethod
    def get_splits(acct:Account, period_starts:list, periods:list):
        """
        get the splits for the account and each sub-account and add to periods
        :param          acct: to get splits
        :param period_starts: start date for each period
        :param       periods: fill with splits for each quarter
        """
        SattoLog.print_text("GnucashUtilities.get_splits()", GnucashUtilities.my_color)

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

                split_amount = GnucashUtilities.gnc_numeric_to_python_decimal(split.GetAmount())

                # if the amount is negative this is a credit, else a debit
                debit_credit_offset = 1 if split_amount < ZERO else 0

                # add the debit or credit to the sum, using the offset to get in the right bucket
                period[2 + debit_credit_offset] += split_amount

                # add the debit or credit to the overall total
                period[4] += split_amount

    @staticmethod
    def fill_splits(root_acct:Account, target_path:list, period_starts:list, periods:list) -> str :
        """
        fill the period list for each account
        :param     root_acct: from the Gnucash book
        :param   target_path: account hierarchy from root account to target account
        :param period_starts: start date for each period
        :param       periods: fill with the splits dates and amounts for requested time span
        :return: name of target_acct
        """
        SattoLog.print_text("GnucashUtilities.fill_splits()", GnucashUtilities.my_color)

        account_of_interest = GnucashUtilities.account_from_path(root_acct, target_path)
        acct_name = account_of_interest.GetName()
        SattoLog.print_text("\naccount_of_interest = {}".format(acct_name), GnucashUtilities.my_color)

        # get the split amounts for the parent account
        GnucashUtilities.get_splits(account_of_interest, period_starts, periods)
        descendants = account_of_interest.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            # print("Descendants of {}:".format(account_of_interest.GetName()))
            for subAcct in descendants:
                # print("{} balance = {}".format(subAcct.GetName(), gnc_numeric_to_python_decimal(subAcct.GetBalance())))
                GnucashUtilities.get_splits(subAcct, period_starts, periods)

        GnucashUtilities.csv_write_period_list(periods)
        return acct_name

    @staticmethod
    def csv_write_period_list(periods:list):
        """
        Write out the details of the submitted period list in csv format
        :param periods: dates and amounts for each quarter
        :return: to stdout
        """
        SattoLog.print_text("GnucashUtilities.csv_write_period_list()", GnucashUtilities.my_color)

        # write out the column headers
        csv_writer = csv.writer(stdout)
        # csv_writer.writerow('')
        csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

        # write out the overall totals for the account of interest
        for start_date, end_date, debit_sum, credit_sum, total in periods:
            csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))

# END class GnucashUtilities
