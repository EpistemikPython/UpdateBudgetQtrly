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
__updated__ = '2019-09-08'

from sys import stdout, path
path.append("/home/marksa/dev/git/Python/Utilities/")
from bisect import bisect_right
from math import log10
from copy import copy
import csv
from gnucash import *
from investment import *


# TODO: SHOULD BE class GnucashSession ?? because NEED to store root_account, book, Commod_table, PriceDB etc ??
def __init__(self, p_monrec: InvestmentRecord, p_mode: str, p_gncfile: str, p_debug: bool, p_domain: str,
             p_pdb: GncPriceDB = None, p_book: Book = None, p_root: Account = None,
             p_curr: GncCommodity = None, p_invrec: InvestmentRecord = None) :
    """
    Create and manage a Gnucash session
    """
    self.logger = Gnulog(p_debug)
    # self.monarch_record = p_mrec
    # self.gnucash_record = p_grec
    self.gnc_file = p_gncfile
    # self.mode = p_mode
    # self.domain = p_domain
    self.price_db = p_pdb
    self.book = p_book
    self.root_acct = p_root
    self.currency = p_curr

    self.logger.print_info("class GnucashSession: Runtime = {}\n".format(dt.now().strftime(DATE_STR_FORMAT)), MAGENTA)


def set_gnc_rec(self, p_gncrec: InvestmentRecord) :
    self.gnucash_record = p_gncrec


# noinspection PyUnboundLocalVariable
def prepare_session(self) :
    """
    initialization needed for a Gnucash session
    :return: message
    """
    self.logger.print_info("prepare_session()", BLUE)
    msg = TEST
    try :
        session = Session(self.gnc_file)
        self.book = session.book

        owner = self.monarch_record.get_owner()
        self.logger.print_info("Owner = {}".format(owner), GREEN)
        self.set_gnc_rec(InvestmentRecord(owner))

        self.create_gnucash_info()

        if self.mode == PROD :
            self.logger.print_info("Mode = {}: COMMIT Price DB edits and Save session.".format(self.mode), GREEN)

            if self.domain != TRADE :
                self.price_db.commit_edit()

            # only ONE session save for the entire run
            session.save()

        session.end()
        # session.destroy()

        msg = self.logger.get_log()

    except Exception as se :
        msg = "prepare_session() EXCEPTION!! '{}'".format(repr(se))
        self.logger.print_error(msg)
        if "session" in locals() and session is not None :
            session.end()
            # session.destroy()
        raise se

    return msg


def begin_session(p_filename:str, p_new:bool=False) -> (Session, Account, GncCommodityTable):
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


def gnc_numeric_to_python_decimal(numeric:GncNumeric, logger:SattoLog=None) -> Decimal:
    """
    convert a GncNumeric value to a python Decimal value
    :param numeric: value to convert
    :param  logger: debug printing
    """
    if logger: logger.print_info("gnucash_utilities.gnc_numeric_to_python_decimal()")

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


def get_accounts(ast_parent: Account, asset_acct_name: str, rev_acct: Account, logger:SattoLog=None) -> (Account, Account) :
    """
    Find the proper Asset and Revenue accounts
    :param      ast_parent: Asset account parent
    :param asset_acct_name: Asset account name
    :param        rev_acct: Revenue account
    :param  logger: debug printing
    :return: Gnucash account, Gnucash account
    """
    logger.print_info('get_accounts()')
    asset_parent = ast_parent
    # special locations for Trust Revenue and Asset accounts
    if asset_acct_name == TRUST_AST_ACCT :
        asset_parent = root_acct.lookup_by_name(TRUST)
        logger.print_info("asset_parent = {}".format(asset_parent.GetName()))
        rev_acct = root_acct.lookup_by_name(TRUST_REV_ACCT)
        logger.print_info("MODIFIED rev_acct = {}".format(rev_acct.GetName()))
    # get the asset account
    asset_acct = asset_parent.lookup_by_name(asset_acct_name)
    if asset_acct is None :
        raise Exception("[164] Could NOT find acct '{}' under parent '{}'".format(asset_acct_name, asset_parent.GetName()))

    logger.print_info("asset_acct = {}".format(asset_acct.GetName()))
    return asset_acct, rev_acct


def get_account_balance(acct:Account, p_date:date, p_currency:GncCommodity, logger:SattoLog=None) -> Decimal:
    """
    get the BALANCE in this account on this date in this currency
    :param       acct: Gnucash Account
    :param     p_date: required
    :param p_currency: Gnucash commodity
    :param     logger: debug printing
    :return: Decimal with balance
    """
    if logger: logger.print_info("gnucash_utilities.get_account_balance()")

    # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
    p_date += ONE_DAY

    acct_bal = acct.GetBalanceAsOfDate(p_date)
    acct_comm = acct.GetCommodity()
    # check if account is already in the desired currency and convert if necessary
    acct_cur = acct_bal if acct_comm == p_currency \
                        else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, p_currency, p_date)

    return gnc_numeric_to_python_decimal(acct_cur)


def get_total_balance(p_root:Account, p_path:list, p_date:date, p_currency:GncCommodity, logger:SattoLog=None) -> (str, Decimal):
    """
    get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
    :param  p_root: Gnucash Account from the Gnucash book
    :param     p_path: path to the account
    :param     p_date: to get the balance
    :param p_currency: Gnucash Commodity: currency to use for the totals
    :param     logger: debug printing
    :return: string, int: account name and account sum
    """
    if logger: logger.print_info("gnucash_utilities.get_total_balance()")

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

    if logger: logger.print_info("gnucash_utilities.get_total_balance(): {} on {} = ${}"
                                 .format(acct_name, p_date, acct_sum))
    return acct_name, acct_sum


def get_account_assets(p_root:Account, asset_accts:dict, end_date:date, p_currency:GncCommodity, logger:SattoLog=None) -> dict:
    """
    Get ASSET data for the specified account for the specified quarter
    :param  asset_accts:
    :param p_root: Gnucash Account from the Gnucash book
    :param     end_date: read the account total at the end of the quarter
    :param   p_currency: Gnucash Commodity: currency to use for the totals
    :param       logger: debug printing
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
        if logger: logger.print_info("GNCU.get_account_assets():\nAssets for {} on {} = ${}\n"
                                     .format(acct_name, end_date, str_sum))
        data[item] = str_sum

    return data


def get_asset_revenue_info(plan_type: str, logger:SattoLog=None) -> (Account, Account) :
    """
    Get the required asset and/or revenue information from each plan
    :param plan_type: plan names from Configuration.InvestmentRecord
    :param    logger: debug printing
    :return: Gnucash account, Gnucash account: revenue account and asset parent account
    """
    logger.print_info("gnucash_utilities.get_asset_revenue_info()")
    rev_path = copy(ACCT_PATHS[REVENUE])
    rev_path.append(plan_type)
    ast_parent_path = copy(ACCT_PATHS[ASSET])
    ast_parent_path.append(plan_type)

    pl_owner = gnucash_record.get_owner()
    if plan_type != OPEN :
        if pl_owner == '' :
            raise Exception("PROBLEM[355]!! Trying to process plan type '{}' but NO Owner value found"
                            " in Tx Collection!!".format(plan_type))
        rev_path.append(ACCT_PATHS[pl_owner])
        ast_parent_path.append(ACCT_PATHS[pl_owner])
    logger.print_info("rev_path = {}".format(str(rev_path)))

    rev_acct = account_from_path(root_acct, rev_path)
    logger.print_info("rev_acct = {}".format(rev_acct.GetName()))
    logger.print_info("asset_parent_path = {}".format(str(ast_parent_path)))
    asset_parent = account_from_path(root_acct, ast_parent_path)
    logger.print_info("asset_parent = {}".format(asset_parent.GetName()))

    return asset_parent, rev_acct


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def account_from_path(top_account:Account, account_path:list, original_path:list=None, logger:SattoLog=None) -> Account:
    """
    RECURSIVE function to get a Gnucash Account: starting from the top account and following the path
    :param   top_account: base Account
    :param  account_path: path to follow
    :param original_path: original call path
    :param        logger: debug printing
    """
    if logger: logger.print_info("gnucash_utilities.account_from_path({}:{})"
                                 .format(top_account.GetName(), account_path), gnc_color)

    if original_path is None:
        original_path = account_path
    account, account_path = account_path[0], account_path[1:]

    account = top_account.lookup_by_name(account)
    if account is None:
        raise Exception("Path '" + str(original_path) + "' could NOT be found!")
    if len(account_path) > 0:
        return account_from_path(account, account_path, original_path)
    else:
        return account


def get_splits(acct:Account, period_starts:list, periods:list, logger:SattoLog=None):
    """
    get the splits for the account and each sub-account and add to periods
    :param          acct: to get splits
    :param period_starts: start date for each period
    :param       periods: fill with splits for each quarter
    :param        logger: debug printing
    """
    if logger: logger.print_info("gnucash_utilities.get_splits()")

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


def fill_splits(p_root:Account, target_path:list, period_starts:list, periods:list, logger:SattoLog=None) -> str :
    """
    fill the period list for each account
    :param     p_root: from the Gnucash book
    :param   target_path: account hierarchy from root account to target account
    :param period_starts: start date for each period
    :param       periods: fill with the splits dates and amounts for requested time span
    :param       logger: debug printing
    :return: name of target_acct
    """
    if logger: logger.print_info("gnucash_utilities.fill_splits()")

    account_of_interest = account_from_path(p_root, target_path)
    acct_name = account_of_interest.GetName()
    if logger: logger.print_info("\naccount_of_interest = {}".format(acct_name))

    # get the split amounts for the parent account
    get_splits(account_of_interest, period_starts, periods)
    descendants = account_of_interest.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        for subAcct in descendants:
            get_splits(subAcct, period_starts, periods)

    csv_write_period_list(periods)
    return acct_name


def create_gnc_price_txs(mtx:dict, ast_parent:Account, rev_acct:Account, logger:SattoLog=None) :
    """
    Create and load Gnucash prices to the Gnucash PriceDB
    :param        mtx: InvestmentRecord transaction
    :param ast_parent: Asset parent account
    :param   rev_acct: Revenue account
    :param     logger: debug printing
    :return: nil
    """
    logger.print_info('create_gnc_price_txs()')
    conv_date = dt.strptime(mtx[DATE], "%d-%b-%Y")
    pr_date = dt(conv_date.year, conv_date.month, conv_date.day)
    datestring = pr_date.strftime("%Y-%m-%d")

    fund_name = mtx[FUND]
    if fund_name in MONEY_MKT_FUNDS:
        return

    int_price = int(mtx[PRICE].replace('.','').replace('$',''))
    val = GncNumeric(int_price, 10000)
    logger.print_info("Adding: {}[{}] @ ${}".format(fund_name, datestring, val))

    pr1 = GncPrice(book)
    pr1.begin_edit()
    pr1.set_time64(pr_date)

    asset_acct, rev_acct = get_accounts(ast_parent, fund_name, rev_acct)
    comm = asset_acct.GetCommodity()
    logger.print_info("Commodity = {}:{}".format(comm.get_namespace(), comm.get_printname()))
    pr1.set_commodity(comm)

    pr1.set_currency(currency)
    pr1.set_value(val)
    pr1.set_source_string("user:price")
    pr1.set_typestr('nav')
    pr1.commit_edit()

    if mode == PROD:
        logger.print_info("Mode = {}: Add Price to DB.".format(self.mode), GREEN)
        price_db.add_price(pr1)
    else:
        logger.print_info("Mode = {}: ABANDON Prices!\n".format(self.mode), RED)


def create_gnc_trade_txs(tx1:dict, tx2:dict, logger:SattoLog=None) :
    """
    Create and load Gnucash transactions to the Gnucash file
    :param    tx1: first transaction
    :param    tx2: matching transaction if a switch
    :param logger: debug printing
    :return: nil
    """
    logger.print_info('create_gnc_trade_txs()')
    # create a gnucash Tx
    gtx = Transaction(book)
    # gets a guid on construction

    gtx.BeginEdit()

    gtx.SetCurrency(currency)
    gtx.SetDate(tx1[TRADE_DAY], tx1[TRADE_MTH], tx1[TRADE_YR])
    # self.dbg.print_info("gtx date = {}".format(gtx.GetDate()), BLUE)
    logger.print_info("tx1[DESC] = {}".format(tx1[DESC]))
    gtx.SetDescription(tx1[DESC])

    # create the ASSET split for the Tx
    spl_ast = Split(book)
    spl_ast.SetParent(gtx)
    # set the account, value, and units of the Asset split
    spl_ast.SetAccount(tx1[ACCT])
    spl_ast.SetValue(GncNumeric(tx1[GROSS], 100))
    spl_ast.SetAmount(GncNumeric(tx1[UNITS], 10000))

    if tx1[SWITCH]:
        # create the second ASSET split for the Tx
        spl_ast2 = Split(book)
        spl_ast2.SetParent(gtx)
        # set the Account, Value, and Units of the second ASSET split
        spl_ast2.SetAccount(tx2[ACCT])
        spl_ast2.SetValue(GncNumeric(tx2[GROSS], 100))
        spl_ast2.SetAmount(GncNumeric(tx2[UNITS], 10000))
        # set Actions for the splits
        spl_ast2.SetAction("Buy" if tx1[UNITS] < 0 else "Sell")
        spl_ast.SetAction("Buy" if tx1[UNITS] > 0 else "Sell")
        # combine Notes for the Tx and set Memos for the splits
        gtx.SetNotes(tx1[NOTES] + " | " + tx2[NOTES])
        spl_ast.SetMemo(tx1[NOTES])
        spl_ast2.SetMemo(tx2[NOTES])
    else:
        # the second split is for a REVENUE account
        spl_rev = Split(book)
        spl_rev.SetParent(gtx)
        # set the Account, Value and Reconciled of the REVENUE split
        spl_rev.SetAccount(tx1[REVENUE])
        rev_gross = tx1[GROSS] * -1
        # self.dbg.print_info("revenue gross = {}".format(rev_gross))
        spl_rev.SetValue(GncNumeric(rev_gross, 100))
        spl_rev.SetReconcile(CREC)
        # set Notes for the Tx
        gtx.SetNotes(tx1[NOTES])
        # set Action for the ASSET split
        action = FEE if FEE in tx1[DESC] else ("Sell" if tx1[UNITS] < 0 else DIST)
        logger.print_info("action = {}".format(action))
        spl_ast.SetAction(action)

    # ROLL BACK if something went wrong and the two splits DO NOT balance
    if not gtx.GetImbalanceValue().zero_p():
        logger.print_error("Gnc tx IMBALANCE = {}!! Roll back transaction changes!"
                                .format(gtx.GetImbalanceValue().to_string()))
        gtx.RollbackEdit()
        return

    if mode == PROD:
        logger.print_info("Mode = {}: Commit transaction changes.\n".format(mode), GREEN)
        gtx.CommitEdit()
    else:
        logger.print_info("Mode = {}: Roll back transaction changes!\n".format(mode), RED)
        gtx.RollbackEdit()


def csv_write_period_list(periods:list, logger:SattoLog=None):
    """
    Write out the details of the submitted period list in csv format
    :param periods: dates and amounts for each quarter
    :param logger: debug printing
    :return: to stdout
    """
    if logger: logger.print_info("gnucash_utilities.csv_write_period_list()")

    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in periods:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))


def show_account(p_root:Account, p_path:list, logger:SattoLog, p_color:str=''):
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
    if len(descendants) == 0:
        logger.print_info("{} has NO Descendants!".format(acct_name), p_color)
    else:
        logger.print_info("Descendants of {}:".format(acct_name), p_color)
