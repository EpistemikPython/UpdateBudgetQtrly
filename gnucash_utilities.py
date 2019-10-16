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
__updated__ = '2019-10-14'

from sys import stdout, path
path.append("/home/marksa/dev/git/Python/Utilities/")
from bisect import bisect_right
from math import log10
from copy import copy
import csv
from gnucash.gnucash_core_c import CREC
from gnucash import *
from investment import *


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

    assert ((10**exponent) == denominator)
    return Decimal((sign, digit_tuple, -exponent))


def csv_write_period_list(periods:list, logger:SattoLog=None):
    """
    Write out the details of the submitted period list in csv format
    :param periods: dates and amounts for each quarter
    :param logger: debug printing
    :return: to stdout
    """
    logger.print_info("gnucash_utilities.csv_write_period_list()")

    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in periods:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))


class GnucashSession:
    def __init__(self, p_mode:str, p_gncfile:str, p_debug:bool, p_domain:str, p_currency:GncCommodity=None) :
        """
        Create and manage a Gnucash session
        init = prepare session, RO or RW, from Gnucash file, debug,
        fxns
            get:
                account(s) of various types
                balances from account(s)
                splits
            create:
                trade txs
                price txs
        """
        self._logger = SattoLog(my_color=GREEN, do_logging=p_debug)
        self._log("\nclass GnucashSession: Runtime = {}\n".format(dt.now().strftime(DATE_STR_FORMAT)), MAGENTA)

        self._gnc_file = p_gncfile
        self._mode     = p_mode
        self._domain   = p_domain

        self._currency = None
        self.set_currency(p_currency)

        self._session      = None
        self._price_db     = None
        self._book         = None
        self._root_acct    = None
        self._commod_table = None

    def _log(self, p_msg:str, p_color:str=''):
        if self._logger:
            self._logger.print_info(p_msg, p_color, p_frame=inspect.currentframe().f_back)

    def _err(self, p_msg:str, err_frame:FrameType):
        if self._logger:
            self._logger.print_info(p_msg, BR_RED, p_frame=err_frame)

    def get_domain(self) -> str:
        return self._domain

    def get_root_acct(self) -> Account:
        return self._root_acct

    def set_currency(self, p_curr:GncCommodity):
        if not p_curr:
            self._log("NO currency!")
            return
        if isinstance(p_curr, GncCommodity):
            self._currency = p_curr
        else:
            self._log("BAD currency '{}' of type: {}".format(str(p_curr), type(p_curr)))

    def begin_session(self, p_new:bool=False):
        self._log('GnucashSession.begin_session()')
        self._session = Session(self._gnc_file, is_new=p_new)
        self._book = self._session.book
        self._root_acct = self._book.get_root_account()
        self._root_acct.get_instance()
        self._commod_table = self._book.get_table()
        if self._currency is None:
            self.set_currency(self._commod_table.lookup("ISO4217", "CAD"))

        if self._domain in (PRICE, BOTH):
            self._price_db = self._book.get_price_db()
            self._price_db.begin_edit()
            self._log("self.price_db.begin_edit()", CYAN)

    def end_session(self, p_save:bool):
        self._log('GnucashSession.end_session()')

        if p_save:
            self._log("Mode = {}: SAVE session.".format(self._mode))
            if self._domain in (PRICE, BOTH):
                self._log("Domain = {}: COMMIT Price DB edits.".format(self._domain))
                self._price_db.commit_edit()
            self._session.save()

        self._session.end()
        # not needed?
        # p_session.destroy()

    def check_end_session(self, p_locals:dict):
        if "gnucash_session" in p_locals and self._session is not None:
            self._session.end()

    def get_account(self, acct_parent:Account, acct_name:str) -> Account:
        """Find the proper account"""
        self._log('GnucashSession.get_account()')

        # special locations for Trust accounts
        if acct_name == TRUST_AST_ACCT :
            return self._root_acct.lookup_by_name(TRUST).lookup_by_name(TRUST_AST_ACCT)
        if acct_name == TRUST_EQY_ACCT:
            return self._root_acct.lookup_by_name(EQUITY).lookup_by_name(TRUST_EQY_ACCT)

        found_acct = acct_parent.lookup_by_name(acct_name)
        if not found_acct:
            raise Exception("Could NOT find acct '{}' under parent '{}'".format(acct_name, acct_parent.GetName()))

        self._log("GnucashSession.get_account() found account: {}".format(found_acct.GetName()))
        return found_acct

    def get_accounts(self, ast_parent:Account, asset_acct_name:str, rev_acct:Account) -> (Account, Account):
        """
        Find the proper Asset and Revenue accounts
        :param      ast_parent: Asset account parent
        :param asset_acct_name: Asset account name
        :param        rev_acct: Revenue account
        :return: Gnucash account, Gnucash account
        """
        self._log('GnucashSession.get_accounts()')
        asset_parent = ast_parent
        # special locations for Trust Revenue and Asset accounts
        if asset_acct_name == TRUST_AST_ACCT :
            asset_parent = self._root_acct.lookup_by_name(TRUST)
            self._log("asset_parent = {}".format(asset_parent.GetName()))
            rev_acct = self._root_acct.lookup_by_name(TRUST_EQY_ACCT)
            self._log("MODIFIED rev_acct = {}".format(rev_acct.GetName()))
        # get the asset account
        asset_acct = asset_parent.lookup_by_name(asset_acct_name)
        if asset_acct is None :
            raise Exception("Could NOT find acct '{}' under parent '{}'".format(asset_acct_name, asset_parent.GetName()))

        self._log("asset_acct = {}".format(asset_acct.GetName()))
        return asset_acct, rev_acct

    def get_account_balance(self, acct:Account, p_date:date, p_currency:GncCommodity=None) -> Decimal:
        """
        get the BALANCE in this account on this date in this currency
        :param       acct: Gnucash Account
        :param     p_date: required
        :param p_currency: Gnucash commodity
        :return: Decimal with balance
        """
        self._log("GnucashSession.get_account_balance()")

        # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
        p_date += ONE_DAY

        currency = self._currency if p_currency is None else p_currency
        acct_bal = acct.GetBalanceAsOfDate(p_date)
        acct_comm = acct.GetCommodity()
        # check if account is already in the desired currency and convert if necessary
        acct_cur = acct_bal if acct_comm == currency \
                            else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, currency, p_date)

        return gnc_numeric_to_python_decimal(acct_cur)

    def get_total_balance(self, p_path:list, p_date:date, p_currency:GncCommodity=None) -> (str, Decimal):
        """
        get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
        :param     p_path: path to the account
        :param     p_date: to get the balance
        :param p_currency: Gnucash Commodity: currency to use for the totals
        :return: string, int: account name and account sum
        """
        currency = self._currency if p_currency is None else p_currency
        acct = self.account_from_path(self._root_acct, p_path)
        acct_name = acct.GetName()
        # get the split amounts for the parent account
        acct_sum = self.get_account_balance(acct, p_date, currency)
        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += self.get_account_balance(sub_acct, p_date, currency)

        self._log("GnucashSession.get_total_balance(): {} on {} = ${}".format(acct_name, p_date, acct_sum))
        return acct_name, acct_sum

    def get_account_assets(self, asset_accts:dict, end_date:date, p_currency:GncCommodity=None) -> dict:
        """
        Get ASSET data for the specified account for the specified quarter
        :param  asset_accts:
        :param     end_date: read the account total at the end of the quarter
        :param   p_currency: Gnucash Commodity: currency to use for the totals
        :return: string with sum of totals
        """
        data = {}
        currency = self._currency if p_currency is None else p_currency
        for item in asset_accts:
            acct_path = asset_accts[item]
            acct = self.account_from_path(self._root_acct, acct_path)
            acct_name = acct.GetName()

            # get the split amounts for the parent account
            acct_sum = self.get_account_balance(acct, end_date, currency)
            descendants = acct.get_descendants()
            if len(descendants) > 0:
                # for EACH sub-account add to the overall total
                # print_info("Descendants of {}:".format(acct_name))
                for sub_acct in descendants:
                    # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                    acct_sum += self.get_account_balance(sub_acct, end_date, currency)

            str_sum = acct_sum.to_eng_string()
            self._log("GnucashSession.get_account_assets():\nAssets for {} on {} = ${}\n"
                      .format(acct_name, end_date, str_sum))
            data[item] = str_sum

        return data

    def get_asset_revenue_info(self, plan_type:str, pl_owner:str) -> (Account, Account):
        """
        Get the required asset and/or revenue information from each plan
        :param plan_type: plan names from Configuration.InvestmentRecord
        :param  pl_owner: needed to find proper revenue account for RRSP & TFSA
        :return: Gnucash account, Gnucash account: revenue account and asset parent account
        """
        self._log("GnucashSession.get_asset_revenue_info()")
        rev_path = copy(ACCT_PATHS[REVENUE])
        rev_path.append(plan_type)
        ast_parent_path = copy(ACCT_PATHS[ASSET])
        ast_parent_path.append(plan_type)

        if plan_type != OPEN :
            if pl_owner == '' :
                raise Exception("Trying to process plan type '{}' but NO Owner value found in Tx Collection!!"
                                .format(plan_type))
            rev_path.append(ACCT_PATHS[pl_owner])
            ast_parent_path.append(ACCT_PATHS[pl_owner])
        self._log("rev_path = {}".format(str(rev_path)))

        rev_acct = self.account_from_path(self._root_acct, rev_path)
        self._log("rev_acct = {}".format(rev_acct.GetName()))
        self._log("asset_parent_path = {}".format(str(ast_parent_path)))
        asset_parent = self.account_from_path(self._root_acct, ast_parent_path)
        self._log("asset_parent = {}".format(asset_parent.GetName()))

        return asset_parent, rev_acct

    def account_from_path(self, top_account:Account, account_path:list) -> Account:
        """
        RECURSIVE function to get a Gnucash Account: starting from the top account and following the path
        :param   top_account: base Account
        :param  account_path: path to follow
        """
        self._log("GnucashSession.account_from_path({}:{})".format(top_account.GetName(), account_path))

        acct_str, acct_path = account_path[0], account_path[1:]

        acct = top_account.lookup_by_name(acct_str)
        if acct is None:
            raise Exception("Path '" + str(account_path) + "' could NOT be found!")
        if len(acct_path) > 0:
            return self.account_from_path(acct, acct_path)
        else:
            return acct

    def show_account(self, p_path:list):
        """
        display an account and its descendants
        :param  p_path: to the account
        :return: nil
        """
        acct = self.account_from_path(self._root_acct, p_path)
        acct_name = acct.GetName()
        self._log("account = {}".format(acct_name))
        descendants = acct.get_descendants()
        if len(descendants) == 0:
            self._log("{} has NO Descendants!".format(acct_name))
        else:
            self._log("Descendants of {}:".format(acct_name))

    def get_splits(self, acct:Account, period_starts:list, periods:list):
        """
        get the splits for the account and each sub-account and add to periods
        :param          acct: to get splits
        :param period_starts: start date for each period
        :param       periods: fill with splits for each quarter
        """
        self._log("GnucashSession.get_splits()")

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

    def fill_splits(self, target_path:list, period_starts:list, periods:list) -> str:
        """
        fill the period list for each account
        :param   target_path: account hierarchy from root account to target account
        :param period_starts: start date for each period
        :param       periods: fill with the splits dates and amounts for requested time span
        :return: name of target_acct
        """
        self._log("GnucashSession.fill_splits()")

        account_of_interest = self.account_from_path(self._root_acct, target_path)
        acct_name = account_of_interest.GetName()
        self._log("\naccount_of_interest = {}".format(acct_name))

        # get the split amounts for the parent account
        self.get_splits(account_of_interest, period_starts, periods)
        descendants = account_of_interest.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            for subAcct in descendants:
                self.get_splits(subAcct, period_starts, periods)

        csv_write_period_list(periods)
        return acct_name

    def create_price_tx(self, mtx:dict, ast_parent:Account, rev_acct:Account):
        """
        Create a PRICE transaction for the current Gnucash session
        :param        mtx: InvestmentRecord transaction
        :param ast_parent: Asset parent account
        :param   rev_acct: Revenue account
        :return: nil
        """
        self._log('GnucashSession.create_price_tx()')
        conv_date = dt.strptime(mtx[DATE], "%d-%b-%Y")
        pr_date = dt(conv_date.year, conv_date.month, conv_date.day)
        datestring = pr_date.strftime("%Y-%m-%d")

        fund_name = mtx[FUND]
        if fund_name in MONEY_MKT_FUNDS:
            return

        int_price = int(mtx[PRICE].replace('.','').replace('$',''))
        val = GncNumeric(int_price, 10000)
        self._log("Adding: {}[{}] @ ${}".format(fund_name, datestring, val))

        pr1 = GncPrice(self._book)
        pr1.begin_edit()
        pr1.set_time64(pr_date)

        asset_acct, rev_acct = self.get_accounts(ast_parent, fund_name, rev_acct)
        comm = asset_acct.GetCommodity()
        self._log("Commodity = {}:{}".format(comm.get_namespace(), comm.get_printname()))
        pr1.set_commodity(comm)

        pr1.set_currency(self._currency)
        pr1.set_value(val)
        pr1.set_source_string("user:price")
        pr1.set_typestr('nav')
        pr1.commit_edit()

        if self._mode == SEND:
            self._log("Mode = {}: Add Price to DB.".format(self._mode), GREEN)
            self._price_db.add_price(pr1)
        else:
            self._log("Mode = {}: ABANDON Prices!\n".format(self._mode), RED)

    def create_trade_tx(self, tx1:dict, tx2:dict):
        """
        Create a TRADE transaction for the current Gnucash session
        :param tx1: first transaction
        :param tx2: matching transaction if a switch
        """
        self._log('GnucashSession.create_trade_tx()')
        # create a gnucash Tx
        gtx = Transaction(self._book)
        # gets a guid on construction

        gtx.BeginEdit()

        gtx.SetCurrency(self._currency)
        gtx.SetDate(tx1[TRADE_DAY], tx1[TRADE_MTH], tx1[TRADE_YR])
        # self.dbg.print_info("gtx date = {}".format(gtx.GetDate()), BLUE)
        self._log("tx1[DESC] = {}".format(tx1[DESC]))
        gtx.SetDescription(tx1[DESC])

        # create the ASSET split for the Tx
        spl_ast = Split(self._book)
        spl_ast.SetParent(gtx)
        # set the account, value, and units of the Asset split
        spl_ast.SetAccount(tx1[ACCT])
        spl_ast.SetValue(GncNumeric(tx1[GROSS], 100))
        spl_ast.SetAmount(GncNumeric(tx1[UNITS], 10000))

        if tx1[SWITCH]:
            # create the second ASSET split for the Tx
            spl_ast2 = Split(self._book)
            spl_ast2.SetParent(gtx)
            # set the Account, Value, and Units of the second ASSET split
            spl_ast2.SetAccount(tx2[ACCT])
            spl_ast2.SetValue(GncNumeric(tx2[GROSS], 100))
            spl_ast2.SetAmount(GncNumeric(tx2[UNITS], 10000))
            # set Actions for the splits
            spl_ast2.SetAction(BUY if tx1[UNITS] < 0 else SELL)
            spl_ast.SetAction(BUY if tx1[UNITS] > 0 else SELL)
            # combine Notes for the Tx and set Memos for the splits
            gtx.SetNotes(tx1[NOTES] + " | " + tx2[NOTES])
            spl_ast.SetMemo(tx1[NOTES])
            spl_ast2.SetMemo(tx2[NOTES])
        elif tx1[TYPE] in (RDMPN, PURCH):
            # the second split is for the HOLD account
            # need a third split if there is a difference b/n Gross and Net
            split_hold = Split(self._book)
            split_hold.SetParent(gtx)
            split_hold.SetAccount(self._root_acct.lookup_by_name(HOLD))
            # compare tx1[GROSS] and tx1[NET] to see if different
            # and need a third split for Financial Services expense
            net_amount = tx1[NET]
            self._log("{} gross = {}".format(HOLD, net_amount))
            if net_amount != tx1[GROSS]:
                split_diff = net_amount - tx1[GROSS]
                split_fin_serv = Split(self._book)
                split_fin_serv.SetParent(gtx)
                split_fin_serv.SetAccount(self._root_acct.lookup_by_name(FIN_SERV))
                split_fin_serv.SetValue(GncNumeric(split_diff, 100))
            split_hold.SetValue(GncNumeric(net_amount * -1, 100))
            gtx.SetNotes(tx1[FUND] + " " + tx1[TYPE])
            # set Action for the ASSET split
            action = SELL if tx1[TYPE] == RDMPN else BUY
            self._log("action = {}".format(action))
            spl_ast.SetAction(action)
        else:
            # the second split is for a REVENUE account
            spl_rev = Split(self._book)
            spl_rev.SetParent(gtx)
            # set the Account, Value and Reconciled of the REVENUE split
            spl_rev.SetAccount(tx1[REVENUE])
            gross_revenue = tx1[GROSS] * -1
            # self.dbg.print_info("gross revenue = {}".format(rev_gross))
            spl_rev.SetValue(GncNumeric(gross_revenue, 100))
            spl_rev.SetReconcile(CREC)
            # set Notes for the Tx
            gtx.SetNotes(tx1[NOTES])
            # set Action for the ASSET split
            action = FEE if FEE in tx1[DESC] else (SELL if tx1[UNITS] < 0 else DIST)
            self._log("action = {}".format(action))
            spl_ast.SetAction(action)

        # ROLL BACK if something went wrong and the splits DO NOT balance
        if not gtx.GetImbalanceValue().zero_p():
            self._err("Gnc tx IMBALANCE = {}!! Roll back transaction changes!"
                      .format(gtx.GetImbalanceValue().to_string()), inspect.currentframe().f_back)
            gtx.RollbackEdit()
            return

        if self._mode == SEND:
            self._log("Mode = {}: Commit transaction changes.\n".format(self._mode), GREEN)
            gtx.CommitEdit()
        else:
            self._log("Mode = {}: Roll back transaction changes!\n".format(self._mode), RED)
            gtx.RollbackEdit()

# END class GnucashSession
