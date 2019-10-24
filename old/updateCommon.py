##############################################################################################################################
# coding=utf-8
#
# updateCommon.py -- common methods and variables for updates
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-07'
__updated__ = '2019-08-25'

from sys import stdout
from datetime import date, timedelta, datetime as dt
from bisect import bisect_right
from decimal import Decimal
from math import log10
from types import FrameType
import csv
import json
import pickle
import inspect
import os.path as osp
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from gnucash import GncNumeric, Account, GncCommodity

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

CELL_TIME_STR:str = "%H:%M:%S"
FILE_DATE_STR:str = "%Y-%m-%d"
FILE_TIME_STR:str = "T%H-%M-%S"
today:dt = dt.now()
now:str  = today.strftime(FILE_DATE_STR + FILE_TIME_STR)

DATE_STR_FORMAT = "\u0023%Y-%m-%d\u0025\u0025%H-%M-%S"
dtnow  = dt.now()
strnow = dtnow.strftime(DATE_STR_FORMAT)

COLOR_FLAG:str = '\x1b['
RED:str     = COLOR_FLAG + '31m'
GREEN:str   = COLOR_FLAG + '32m'
BROWN:str   = COLOR_FLAG + '33m'
BLUE:str    = COLOR_FLAG + '34m'
MAGENTA:str = COLOR_FLAG + '35m'
CYAN:str    = COLOR_FLAG + '36m'
BR_RED:str   = COLOR_FLAG + '91m'
BR_GREEN:str = COLOR_FLAG + '92m'
YELLOW:str   = COLOR_FLAG + '93m'
BR_BLUE:str  = COLOR_FLAG + '94m'
PINK:str     = COLOR_FLAG + '95m'
BR_CYAN:str  = COLOR_FLAG + '96m'
BLACK:str     = COLOR_FLAG + '30m'
GREY:str      = COLOR_FLAG + '90m'
WHITE:str     = COLOR_FLAG + '37m'
COLOR_OFF:str = COLOR_FLAG + '0m'


class SattoLog:
    def __init__(self, p_debug:bool):
        self.debug = p_debug
        self.log_text = []

    def append(self, obj:object):
        self.log_text.append(obj)

    def clear_log(self):
        self.log_text = []

    def get_log(self) -> list :
        return self.log_text

    def print_info(self, info:object, color:str='', inspector:bool=True, newline:bool=True) -> str :
        """
        Print information with choices of color, inspection info, newline
        """
        text = str(info)
        if self.debug:
            calling_frame = inspect.currentframe().f_back
            # print(inspect.getframeinfo(calling_frame))
            # print("frame type = {}, class = {}, module = {}, class module = {}, package = {}"
            #       .format(type(calling_frame), calling_frame.__class__, inspect.getmodule(calling_frame),
            #               'calling_frame.__class__.__module__', 'calling_frame.__package__'))
            self.print_text(info, color, inspector, newline, calling_frame)
            self.append(text)
        return text

    def print_error(self, info:object) -> str :
        """
        Print Error information in RED with inspection info
        """
        text = str(info)
        if self.debug:
            calling_frame = inspect.currentframe().f_back
            self.print_warning(info, calling_frame)
        return text

    @staticmethod
    def print_warning(info:object, p_frame:FrameType=None) -> str :
        calling_frame = p_frame if p_frame is not None else inspect.currentframe().f_back
        return SattoLog.print_text(info, BR_RED, p_frame=calling_frame)

    @staticmethod
    def print_text(info:object, color:str='', inspector:bool=True, newline:bool=True, p_frame:FrameType=None) -> str :
        """
        Print information with choices of color, inspection info, newline
        """
        inspect_line = ''
        if info is None:
            info = '==========================================================================================================='
            inspector = False
        text = str(info)
        if inspector:
            calling_frame = p_frame if p_frame is not None else inspect.currentframe().f_back
            parent_frame = calling_frame.f_back
            calling_file  = inspect.getfile(calling_frame).split('/')[-1]
            parent_file  = inspect.getfile(parent_frame).split('/')[-1]
            calling_line  = str(inspect.getlineno(calling_frame))
            parent_line  = str(inspect.getlineno(parent_frame))
            inspect_line  = '[' + parent_file + '@' + parent_line + '->' + calling_file + '@' + calling_line + ']: '
        print(inspect_line + color + text + COLOR_OFF, end=('\n' if newline else ''))
        return text

# END class SattoLog


class CommonUtilities:
    def __init__(self):
        SattoLog.print_text("CommonUtilities", GREEN)

    my_color = BROWN

    @staticmethod
    def year_span(target_year:int, base_year:int, base_year_span:int, hdr_span:int) -> int :
        """
        calculate which row to update, factoring in the header row placed every so-many years
        :param    target_year: year to calculate for
        :param      base_year: starting year
        :param base_year_span: number of rows between equivalent positions in adjacent years, not including header rows
        :param       hdr_span: number of rows between header rows
        """
        SattoLog.print_text("CommonUtilities.year_span()", CommonUtilities.my_color)

        year_diff = int(target_year - base_year)
        hdr_adjustment = 0 if hdr_span <= 0 else (year_diff // int(hdr_span))
        return int(year_diff * base_year_span) + hdr_adjustment

    @staticmethod
    def get_int_year(target_year:str, base_year:int) -> int :
        """
        convert the string representation of a year to an int
        :param target_year: to convert
        :param   base_year: earliest possible year
        """
        SattoLog.print_text("CommonUtilities.get_int_year()", CommonUtilities.my_color)

        if not target_year.isnumeric():
            SattoLog.print_warning("Input MUST be the String representation of a Year, e.g. '2013'!")
            exit(214)
        int_year = int(float(target_year))
        if int_year > today.year or int_year < base_year:
            SattoLog.print_warning("Input MUST be the String representation of a Year between {} and {}!"
                                   .format(today.year, base_year))
            exit(219)

        return int_year

    @staticmethod
    def get_int_quarter(p_qtr:str) -> int :
        """
        convert the string representation of a quarter to an int
        :param  p_qtr: to convert
        """
        SattoLog.print_text("CommonUtilities.get_int_quarter()", CommonUtilities.my_color)

        if not p_qtr.isnumeric():
            SattoLog.print_warning("Input MUST be a String of 0..4!")
            exit(233)
        int_qtr = int(float(p_qtr))
        if int_qtr > 4 or int_qtr < 0:
            SattoLog.print_warning("Input MUST be a String of 0..4!")
            exit(237)

        return int_qtr

    @staticmethod
    def next_quarter_start(start_year:int, start_month:int) -> (int, int) :
        """
        get the year and month that starts the FOLLOWING quarter
        :param  start_year
        :param start_month
        """
        SattoLog.print_text("CommonUtilities.next_quarter_start()", CommonUtilities.my_color)

        # add number of months for a Quarter
        next_month = start_month + PERIOD_QTR

        # use integer division to find out if the new end month is in a different year,
        # what year it is, and what the end month number should be changed to.
        next_year = start_year + ((next_month - 1) // YEAR_MONTHS)
        next_month = ((next_month - 1) % YEAR_MONTHS) + 1

        return next_year, next_month

    @staticmethod
    def current_quarter_end(start_year:int, start_month:int) -> date :
        """
        get the date that ends the CURRENT quarter
        :param  start_year
        :param start_month
        """
        SattoLog.print_text("CommonUtilities.current_quarter_end()", CommonUtilities.my_color)

        end_year, end_month = CommonUtilities.next_quarter_start(start_year, start_month)

        # last step, the end date is one day back from the start of the next period
        # so we get a period end like 2010-03-31 instead of 2010-04-01
        return date(end_year, end_month, 1) - ONE_DAY

    @staticmethod
    def generate_quarter_boundaries(start_year:int, start_month:int, num_qtrs:int) -> (date, date) :
        """
        get the start and end dates for the quarters in the submitted range
        :param  start_year
        :param start_month
        :param    num_qtrs: number of quarters to calculate
        """
        SattoLog.print_text("CommonUtilities.generate_quarter_boundaries()", CommonUtilities.my_color)

        for i in range(num_qtrs):
            yield(date(start_year, start_month, 1), CommonUtilities.current_quarter_end(start_year, start_month))
            start_year, start_month = CommonUtilities.next_quarter_start(start_year, start_month)

    @staticmethod
    def save_to_json(fname:str, ts:str, json_data, indt:int=4) -> str:
        """
        print json data to a file -- add a time string to get a unique file name each run
        :param     fname: file path and name
        :param        ts: timestamp to use
        :param json_data: json compatible struct
        :param      indt: indentation amount
        :return: file name
        """
        SattoLog.print_text("CommonUtilities.save_to_json()", CommonUtilities.my_color)

        out_file = fname + '_' + ts + ".json"
        SattoLog.print_text("json file is '{}'\n".format(out_file), MAGENTA)
        fp = open(out_file, 'w')
        json.dump(json_data, fp, indent=indt)
        fp.close()
        return out_file

# END class CommonUtilities


class GoogleUtilities:
    def __init__(self):
        SattoLog.print_text("GoogleUtilities", GREEN)

    my_color = CYAN

    CREDENTIALS_FILE:str = 'secrets/credentials.json'

    SHEETS_RW_SCOPE:list = ['https://www.googleapis.com/auth/spreadsheets']

    SHEETS_EPISTEMIK_RW_TOKEN:dict = {
        'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
        'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
        'P4' : 'secrets/token.sheets.epistemik.rw.pickle4'
    }
    TOKEN:str = SHEETS_EPISTEMIK_RW_TOKEN['P4']

    # Spreadsheet ID
    BUDGET_QTRLY_ID_FILE:str = 'secrets/Budget-qtrly.id'

    @staticmethod
    def fill_cell(sheet:str, col:str, row:int, val, data:list=None) -> list :
        """
        create a dict of update information for one Google Sheets cell and add to the submitted or created list
        :param sheet:  particular sheet in my Google spreadsheet to update
        :param   col:  column to update
        :param   row:  to update
        :param   val:  str OR Decimal: value to fill with
        :param  data:  optional list to append with created dict
        """
        # SattoLog.print_text("GoogleUtilities.fill_cell()", GoogleUtilities.my_color)

        if data is None:
            data = list()

        value = val.to_eng_string() if isinstance(val, Decimal) else val
        cell = {'range': sheet + '!' + col + str(row), 'values': [[value]]}
        SattoLog.print_text("GoogleUtilities.fill_cell() = {}\n".format(cell), GoogleUtilities.my_color)
        data.append(cell)

        return data

    @staticmethod
    def get_budget_id() -> str :
        """
        get the budget id string from the file in the secrets folder
        """
        SattoLog.print_text("GoogleUtilities.get_budget_id()", GoogleUtilities.my_color)

        fp = open(GoogleUtilities.BUDGET_QTRLY_ID_FILE, "r")
        fid = fp.readline().strip()
        SattoLog.print_text("Budget Id = '{}'\n".format(fid), GoogleUtilities.my_color)
        fp.close()
        return fid

    @staticmethod
    def get_credentials() -> pickle :
        """
        get the proper credentials needed to write to the Google spreadsheet
        """
        SattoLog.print_text("GoogleUtilities.get_credentials()", GoogleUtilities.my_color)

        creds = None
        if osp.exists(GoogleUtilities.TOKEN):
            with open(GoogleUtilities.TOKEN, 'rb') as token:
                creds = pickle.load(token)

        # if there are no (valid) credentials available, let the user log in.
        if creds is None or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(GoogleUtilities.CREDENTIALS_FILE,
                                                                 GoogleUtilities.SHEETS_RW_SCOPE)
                creds = flow.run_local_server()
            # save the credentials for the next run
            with open(GoogleUtilities.TOKEN, 'wb') as token:
                pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

        return creds

    @staticmethod
    def send_data(mode:str, data:list) -> dict :
        """
        Send the data list to the document
        :param mode: '.?[send][1]'
        :param data: Gnucash data for each needed quarter
        :return: server response
        """
        SattoLog.print_text("GoogleUtilities.send_data({})\n".format(mode), GoogleUtilities.my_color)

        response = None
        try:
            assets_body = {
                'valueInputOption': 'USER_ENTERED',
                'data': data
            }

            if 'send' in mode:
                creds = GoogleUtilities.get_credentials()
                service = build('sheets', 'v4', credentials=creds)
                vals = service.spreadsheets().values()
                response = vals.batchUpdate(spreadsheetId=GoogleUtilities.get_budget_id(), body=assets_body).execute()

                SattoLog.print_text('{} cells updated!\n'.format(response.get('totalUpdatedCells')), GoogleUtilities.my_color)

        except Exception as se:
            SattoLog.print_warning("Exception: {}!".format(se))
            exit(419)

        return response

# END class GoogleUtilities


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
    def get_total_balance(root_acct, path: list, p_date: date, p_currency):
        """
        get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
        :param  root_acct: Gnucash Account from the Gnucash book
        :param       path: path to the account
        :param     p_date: to get the balance
        :param p_currency: Gnucash Commodity: currency to use for the totals
        :return: string, int: account name and account sum
        """
        # SattoLog.print_text("GnucashUtilities.get_total_balance()", GnucashUtilities.my_color)

        acct = GnucashUtilities.account_from_path(root_acct, path)
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
