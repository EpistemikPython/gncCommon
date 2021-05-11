##############################################################################################################################
# coding=utf-8
#
# gnucash_utilities.py -- useful classes, functions & constants
#
# some code from gnucash examples by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
#
# Copyright (c) 2019-21 Mark Sattolo <epistemik@gmail.com>

__author__       = "Mark Sattolo"
__author_email__ = "epistemik@gmail.com"
__gnucash_version__ = "?3.5+"
__created__ = "2019-04-07"
__updated__ = "2021-05-11"

import threading
from datetime import date
from sys import stdout, path
from bisect import bisect_right
from math import log10
from copy import copy
import csv
from gnucash import GncNumeric, GncCommodity, GncPrice, Account, Session, Split, Transaction
from gnucash.gnucash_core_c import CREC
from mhsUtils import Decimal, ZERO, ONE_DAY
path.append("/newdata/dev/git/Python/Gnucash/common")
from investment import *


def gnc_numeric_to_python_decimal(numeric:GncNumeric, logger:lg.Logger=None) -> Decimal:
    """
    convert a GncNumeric value to a python Decimal value
    :param numeric: value to convert
    :param  logger
    """
    if logger: logger.debug(F"numeric = {numeric.num()}/{numeric.denom()}")

    negative = numeric.negative_p()
    sign = 1 if negative else 0

    val = GncNumeric(numeric.num(), numeric.denom())
    result = val.to_decimal(None)
    if not result:
        raise Exception(F"GncNumeric value '{val.to_string()}' CANNOT be converted to decimal!")

    digit_tuple = tuple(int(char) for char in str(val.num()) if char != '-')
    denominator = val.denom()
    exponent = int(log10(denominator))

    assert ((10**exponent) == denominator)
    return Decimal((sign, digit_tuple, -exponent))


def get_splits(p_acct:Account, period_starts:list, periods:list, logger:lg.Logger=None):
    """
    get the splits for the account and each sub-account and add to periods
    :param        p_acct: to get splits
    :param period_starts: start date for each period
    :param       periods: fill with splits for each quarter
    :param        logger
    """
    if logger: logger.debug(F"account = {p_acct.GetName()}, period starts = {period_starts}, periods = {periods}")
    # insert and add all splits in the periods of interest
    for split in p_acct.GetSplitList():
        trans = split.parent
        # GetDate() returns a datetime but need a date
        trans_date = trans.GetDate().date()

        # use binary search to find the period that starts before or on the transaction date
        period_index = bisect_right(period_starts, trans_date) - 1

        # ignore transactions with a date before the matching period start and after the last period_end
        if period_index >= 0 and trans_date <= periods[len(periods) - 1][1]:
            # get the period bucket appropriate for the split in question
            period = periods[period_index]
            assert (period[1] >= trans_date >= period[0])

            split_amount = gnc_numeric_to_python_decimal(split.GetAmount())

            # if the amount is negative this is a credit, else a debit
            debit_credit_offset = 1 if split_amount < ZERO else 0

            # add the debit or credit to the sum, using the offset to get in the right bucket
            period[2 + debit_credit_offset] += split_amount

            # add the debit or credit to the overall total
            period[4] += split_amount


def fill_splits(base_acct:Account, target_path:list, period_starts:list, periods:list, logger:lg.Logger=None) -> str:
    """
    fill the period list for each account
    :param     base_acct: base account
    :param   target_path: account hierarchy from base account to target account
    :param period_starts: start date for each period
    :param       periods: fill with the splits dates and amounts for requested time span
    :param        logger
    :return: name of target_acct
    """
    account_of_interest = account_from_path(base_acct, target_path, logger)
    acct_name = account_of_interest.GetName()
    if logger: logger.debug(F"base account = {base_acct.GetName()}; account of interest = {acct_name}")

    # get the split amounts for the parent account
    get_splits(account_of_interest, period_starts, periods, logger)
    descendants = account_of_interest.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        for subAcct in descendants:
            get_splits(subAcct, period_starts, periods)

    if logger and logger.level < lg.DEBUG:
        csv_write_period_list(periods)

    return acct_name


def account_from_path(top_account:Account, account_path:list, logger:lg.Logger=None) -> Account:
    """
    RECURSIVE function to get a Gnucash Account: starting from the top account and following the path
    :param   top_account: base Account
    :param  account_path: path to follow
    :param        logger
    """
    if logger: logger.debug(F"top account = {top_account.GetName()}; account path = {account_path}")

    acct_name = account_path[0]
    acct_path = account_path[1:]

    acct = top_account.lookup_by_name(acct_name)
    if acct is None:
        raise Exception(F"Path '{str(account_path)}' could NOT be found!")
    if len(acct_path) > 0:
        return account_from_path(acct, acct_path)
    else:
        return acct


def csv_write_period_list(periods:list, logger:lg.Logger=None):
    """
    Write out the details of the submitted period list in csv format
    :param   periods: dates and amounts for each quarter
    :param   logger
    :return: to stdout
    """
    if logger: logger.debug(F"periods = {periods}")

    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in periods:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))


class GnucashSession:
    """
    Create, manage and terminate a Gnucash session
    fxns:
        get:
            account(s) of various types
            balances from account(s)
            splits
        create:
            trade txs
            price txs
    """
    # PREVENT multiple instances/threads from trying to use the SAME Gnucash file AT THE SAME TIME
    _lock = dict()

    def __init__(self, p_mode:str, p_gncfile:str, p_domain:str, p_logger:lg.Logger, p_currency:GncCommodity=None):
        self._lgr = p_logger
        self._lgr.info(F"\n\tLaunch {self.__class__.__name__} instance on file {p_gncfile}\n\t"
                       F" at Runtime = {get_current_time()}\n")

        self._gnc_file = p_gncfile
        self._mode     = p_mode   # test or send
        self._domain   = p_domain # txs and/or prices

        self._currency = None
        self.set_currency(p_currency)

        self._session      = None
        self._price_db     = None
        self._book         = None
        self._root_acct    = None
        self._commod_table = None

        if self._gnc_file not in self._lock:
            self._lock[self._gnc_file] = threading.Lock()
        else:
            self._lgr.warning(F"lock {self._gnc_file} ALREADY defined!")
        self._lgr.info(F"locks defined = {str(self._lock)}")

    def get_domain(self) -> str:
        return self._domain

    def get_root_acct(self) -> Account:
        return self._root_acct

    def get_file_name(self):
        return self._gnc_file

    def add_price(self, prc:GncPrice):
        self._price_db.add_price(prc)

    def set_currency(self, p_curr:GncCommodity):
        if not p_curr:
            self._lgr.warning('NO currency!')
            return
        if isinstance(p_curr, GncCommodity):
            self._currency = p_curr
            self._lgr.debug(F"currency set to {p_curr}")
        else:
            self._lgr.error(F"BAD currency '{str(p_curr)}' of type: {type(p_curr)}")

    def begin_session(self, p_new:bool=False):
        # PREVENT being able to start a separate Session with this Gnucash file
        self._lock[self._gnc_file].acquire()
        self._lgr.info(F"acquired lock {self._gnc_file} at {get_current_time()}")

        self._session = Session(self._gnc_file, is_new=p_new)
        self._book = self._session.book
        self._root_acct = self._book.get_root_account()
        self._root_acct.get_instance()
        self._commod_table = self._book.get_table()

        if self._currency is None:
            self.set_currency(self._commod_table.lookup("ISO4217", "CAD"))

        if self._domain in (PRICE,BOTH):
            self._price_db = self._book.get_price_db()
            self._lgr.debug('self.price_db.begin_edit()')
            self._price_db.begin_edit()

    def end_session(self, p_save:bool=None):
        self._lgr.debug(get_current_time())

        save_session = p_save if p_save else (self._mode == SEND)
        if save_session:
            self._lgr.info(F"Mode = {self._mode}: SAVE session.")
            if self._domain in (PRICE,BOTH):
                self._lgr.info(F"Domain = {self._domain}: COMMIT Price DB edits.")
                self._price_db.commit_edit()
            self._session.save()

        self._session.end()

        # RELEASE the thread lock on this Gnucash file
        self._lock[self._gnc_file].release()
        self._lgr.info(F"released lock {self._gnc_file} at {get_current_time()}")

    def check_end_session(self, p_locals:dict):
        if "gnucash_session" in p_locals and self._session is not None:
            self._session.end()

    def get_account(self, acct_name:str, acct_parent:Account=None) -> Account:
        """Under the specified parent account find the Account with the specified name"""
        if acct_parent is None:
            acct_parent = self.get_root_acct()
        acct_parent_name = acct_parent.GetName()
        self._lgr.debug(F"account parent = {acct_parent_name}; account = {acct_name}")

        # special location for Trust account
        if acct_name == TRUST_AST_ACCT:
            found_acct = self._root_acct.lookup_by_name(TRUST).lookup_by_name(TRUST_AST_ACCT)
        # elif acct_name == TRUST_EQY_ACCT:
        #     found_acct = self._root_acct.lookup_by_name(EQUITY).lookup_by_name(TRUST_EQY_ACCT)
        else:
            found_acct = acct_parent.lookup_by_name(acct_name)

        if not found_acct:
            raise Exception(F"Could NOT find acct '{acct_name}' under parent '{acct_parent_name}'")

        self._lgr.debug(F"Found account: {found_acct.GetName()}")
        return found_acct

    def get_account_balance(self, acct:Account, p_date:date, p_currency:GncCommodity=None) -> Decimal:
        """
        get the BALANCE in this account on this date in this currency
        :param       acct: Gnucash Account
        :param     p_date: required
        :param p_currency: Gnucash commodity
        :return: Decimal with balance
        """
        # CALLS ARE RETRIEVING ACCOUNT BALANCES FROM DAY BEFORE!!??
        p_date += ONE_DAY

        currency = self._currency if p_currency is None else p_currency
        acct_bal = acct.GetBalanceAsOfDate(p_date)
        acct_comm = acct.GetCommodity()
        # check if account is already in the desired currency and convert if necessary
        acct_cur = acct_bal if acct_comm == currency \
                            else acct.ConvertBalanceToCurrencyAsOfDate(acct_bal, acct_comm, currency, p_date)

        return gnc_numeric_to_python_decimal(acct_cur)

    def get_total_balance(self, p_path:list, p_date:date, p_currency:GncCommodity=None) -> Decimal:
        """
        get the total BALANCE in the account and all sub-accounts on this path on this date in this currency
        :param     p_path: path to the account
        :param     p_date: to get the balance
        :param p_currency: Gnucash Commodity: currency to use for the totals
        """
        currency = self._currency if p_currency is None else p_currency
        acct = account_from_path(self._root_acct, p_path)
        # get the split amounts for the parent account
        acct_sum = self.get_account_balance(acct, p_date, currency)

        descendants = acct.get_descendants()
        if len(descendants) > 0:
            # for EACH sub-account add to the overall total
            for sub_acct in descendants:
                # ?? GETTING SLIGHT ROUNDING ERRORS WHEN ADDING MUTUAL FUND VALUES...
                acct_sum += self.get_account_balance(sub_acct, p_date, currency)

        self._lgr.debug(F"{acct.GetName()} on {p_date} = {acct_sum}")
        return acct_sum

    def get_account_assets(self, asset_accts:dict, end_date:date, p_currency:GncCommodity=None, p_data:dict=None) -> dict:
        """
        Get ASSET data for the specified accounts for the specified date
        :param       p_data: optional dict for data
        :param  asset_accts: from Gnucash file
        :param     end_date: on which to read the account total
        :param   p_currency: Gnucash Commodity: currency to use for the sums
        :return: dict with amounts
        """
        self._lgr.debug(F"end_date = {end_date}")

        data = {} if p_data is None else p_data
        currency = self._currency if p_currency is None else p_currency

        for item in asset_accts:
            acct_sum = self.get_total_balance(asset_accts[item], end_date, currency)
            data[item] = acct_sum.to_eng_string()

        return data

    def _get_asset_or_revenue_account(self, acct_type:str, plan_type:str, pl_owner:str) -> Account:
        """
        Get the required asset and/or revenue account
        :param acct_type: from Gnucash file
        :param plan_type: plan names from investment.InvestmentRecord
        :param  pl_owner: needed to find proper account for RRSP & TFSA plan types
        """
        self._lgr.debug(F"account type = {acct_type}; plan type = {plan_type}; plan owner = {pl_owner}")

        if acct_type not in (ASSET,REV):
            raise Exception(F"GnucashSession._get_asset_or_revenue_account(): BAD Account type: {acct_type}!")
        account_path = copy(ACCT_PATHS[acct_type])

        if plan_type not in (OPEN,RRSP,TFSA):
            raise Exception(F"GnucashSession._get_asset_or_revenue_account(): BAD Plan type: {plan_type}!")
        account_path.append(plan_type)

        if plan_type in (RRSP,TFSA):
            if pl_owner not in (MON_MARK,MON_LULU):
                raise Exception(F"GnucashSession._get_asset_or_revenue_account(): BAD Owner value: {pl_owner}!")
            account_path.append(ACCT_PATHS[pl_owner])

        target_account = account_from_path(self._root_acct, account_path)
        self._lgr.debug(F"target_account = {target_account.GetName()}")

        return target_account

    def get_asset_account(self, plan_type:str, pl_owner:str) -> Account:
        self._lgr.debug(get_current_time())
        return self._get_asset_or_revenue_account(ASSET, plan_type, pl_owner)

    def get_revenue_account(self, plan_type:str, pl_owner:str) -> Account:
        self._lgr.debug(get_current_time())
        return self._get_asset_or_revenue_account(REV, plan_type, pl_owner)

    def show_account(self, p_path:list):
        """
        display an account and its descendants
        :param  p_path: to the account
        """
        acct = account_from_path(self._root_acct, p_path)
        acct_name = acct.GetName()
        self._lgr.debug(F"account = {acct_name}")

        descendants = acct.get_descendants()
        if len(descendants) == 0:
            self._lgr.debug(F"{acct_name} has NO Descendants!")
        else:
            self._lgr.debug(F"Descendants of {acct_name}:")
            for item in descendants:
                self._lgr.debug(F"account = {item.GetName()}")

    def create_price(self, mtx:dict, ast_parent:Account):
        """
        Create a PRICE DB entry for the current Gnucash session
        :param        mtx: InvestmentRecord transaction
        :param ast_parent: Asset parent account
        """
        self._lgr.debug(F"asset parent = {ast_parent}")

        conv_date = dt.strptime(mtx[DATE], "%d-%b-%Y")
        pr_date = dt(conv_date.year, conv_date.month, conv_date.day)
        datestring = pr_date.strftime("%Y-%m-%d")

        fund_name = mtx[FUND]
        if fund_name in MONEY_MKT_FUNDS:
            return

        int_price = int(mtx[PRICE].replace('.','').replace('$',''))
        val = GncNumeric(int_price, 10000)
        self._lgr.debug(F"Adding: {fund_name}[{datestring}] @ ${val}")

        gnc_price = GncPrice(self._book)
        gnc_price.begin_edit()
        gnc_price.set_time64(pr_date)

        asset_acct = self.get_account(fund_name, ast_parent)
        comm = asset_acct.GetCommodity()
        self._lgr.debug(F"Commodity = {comm.get_namespace()}:{comm.get_printname()}")
        gnc_price.set_commodity(comm)

        gnc_price.set_currency(self._currency)
        gnc_price.set_value(val)
        gnc_price.set_source_string('user:price')
        gnc_price.set_typestr('nav')
        gnc_price.commit_edit()

        if self._mode == SEND:
            self._lgr.debug(F"Mode = {self._mode}: Add Price to DB.")
            self.add_price(gnc_price)
        else:
            self._lgr.warning(F"Mode = {self._mode}: ABANDON Prices!\n")

    def create_trade_tx(self, tx1:dict, tx2:dict):
        """
        Create a TRADE transaction for the current Gnucash session
        :param tx1: first transaction
        :param tx2: matching transaction if a switch
        """
        self._lgr.debug(get_current_time())

        # create a gnucash Tx -- gets a guid on construction
        gtx = Transaction(self._book)

        gtx.BeginEdit()

        gtx.SetCurrency(self._currency)
        gtx.SetDate(tx1[TRADE_DAY], tx1[TRADE_MTH], tx1[TRADE_YR])
        self._lgr.debug(F"tx1[DESC] = {tx1[DESC]}")
        gtx.SetDescription(tx1[DESC])

        # create the ASSET split for the Tx
        spl_ast = Split(self._book)
        spl_ast.SetParent(gtx)
        # set the account, value, and units of the Asset split
        spl_ast.SetAccount(tx1[ACCT])
        spl_ast.SetValue(GncNumeric(tx1[GROSS], 100))
        spl_ast.SetAmount(GncNumeric(tx1[UNITS], 10000))

        # create the second split for the Tx
        split_2 = Split(self._book)
        split_2.SetParent(gtx)

        if tx1[TYPE] in PAIRED_TYPES:
            # the second split is also an ASSET
            self._lgr.debug(F"tx2[DESC] = {tx2[DESC]}")
            split_2.SetAccount(tx2[ACCT])
            split_2.SetValue(GncNumeric(tx2[GROSS], 100))
            split_2.SetAmount(GncNumeric(tx2[UNITS], 10000))
            # set Actions for the splits
            split_2.SetAction(BUY if tx1[UNITS] < 0 else SELL)
            spl_ast.SetAction(BUY if tx1[UNITS] > 0 else SELL)
            # modify the Gnc Tx description to note the paired account
            gtx.SetDescription(tx1[DESC] + ' <> ' + tx2[FUND])
            gtx.SetNotes(tx1[NOTES] + " | " + tx2[NOTES])
            spl_ast.SetMemo(tx1[NOTES])
            split_2.SetMemo(tx2[NOTES])
        elif tx1[TYPE] in (RDMPN,PURCH):
            # the second split is for the HOLD account
            split_2.SetAccount(self._root_acct.lookup_by_name(HOLD))
            # MAY need a THIRD split for Financial Services expense e.g. fees, commissions
            # compare tx1[GROSS] and tx1[NET]
            if tx1[NET] != tx1[GROSS]:
                self._lgr.debug(F"Tx net '{tx1[NET]}' != gross '{tx1[GROSS]}'")
                amount_diff = tx1[NET] - tx1[GROSS]
                split_fin_serv = Split(self._book)
                split_fin_serv.SetParent(gtx)
                split_fin_serv.SetAccount(self._root_acct.lookup_by_name(FIN_SERV))
                split_fin_serv.SetValue(GncNumeric(amount_diff, 100))
            split_2.SetValue(GncNumeric(tx1[NET] * -1, 100))
            gtx.SetNotes(tx1[TYPE] + ": " + tx1[NOTES] if tx1[NOTES] else tx1[FUND])
            # set Action for the ASSET split
            action = SELL if tx1[TYPE] == RDMPN else BUY
            spl_ast.SetAction(action)
        else:
            # the second split is for a REVENUE account e.g. 'Investment Income'
            split_2.SetAccount(tx1[REV])
            gross_revenue = tx1[GROSS] * -1
            split_2.SetValue(GncNumeric(gross_revenue, 100))
            split_2.SetReconcile(CREC)
            gtx.SetNotes(tx1[NOTES])
            # set Action for the ASSET split
            action = FEE if FEE in tx1[DESC] else (SELL if tx1[UNITS] < 0 else DIST)
            spl_ast.SetAction(action)

        # ROLL BACK if something went wrong and the splits DO NOT balance
        if not gtx.GetImbalanceValue().zero_p():
            self._lgr.error(F"Gnc tx IMBALANCE = {gtx.GetImbalanceValue().to_string()}!! Roll back transaction changes!")
            gtx.RollbackEdit()
            return

        if self._mode == SEND:
            self._lgr.info(F"Mode = {self._mode}: Commit transaction.")
            gtx.CommitEdit()
        else:
            self._lgr.warning(F"Mode = {self._mode}: ROLL BACK transaction!\n")
            gtx.RollbackEdit()

# END class GnucashSession
