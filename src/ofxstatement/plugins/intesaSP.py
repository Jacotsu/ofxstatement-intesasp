#!/usr/bin/env python3.6

from openpyxl import load_workbook
from dataclasses import dataclass
from datetime import datetime
import logging

from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import (Statement, StatementLine,
                                    generate_transaction_id)

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger('IntesaSP')


@dataclass
class Movimento:
    data_contabile: datetime
    data_valuta: datetime
    descrizione: str
    accrediti: float
    addebiti: float
    descrizione_estesa: str
    mezzo: str


class IntesaSanPaoloPlugin(Plugin):

    def get_parser(self, filename):
        parser = IntesaSanPaoloXlsxParser(filename)
        parser.statement.bank_id = self.settings.get('BIC', 'IntesaSP')
        return parser


class IntesaSanPaoloXlsxParser(StatementParser):

    def split_records(self):
        return self._get_movimenti()

    def __init__(self, filename):
        self.fin = filename
        self.statement = Statement()
        self.statement.account_id = self._get_account_id()
        self.statement.currency = self._get_currency()
        self.statement.start_balance = self._get_start_balance()
        self.statement.start_date = self._get_start_date()
        self.statement.end_balance = self._get_end_balance()
        self.statement.end_date = self._get_end_date()
        logging.debug(self.statement)

    def parse(self):
        return super(IntesaSanPaoloXlsxParser, self).parse()

    def parse_record(self, mov):
        stat_line = StatementLine(None,
                                  mov.data_contabile,
                                  mov.descrizione_estesa,
                                  mov.accrediti if mov.accrediti else
                                  mov.addebiti)
        stat_line.id = generate_transaction_id(stat_line)
        stat_line.date_user = mov.data_valuta
        stat_line.trntype = IntesaSanPaoloXlsxParser._get_transaction_type(mov)
        logging.debug(stat_line)
        return stat_line

    def _get_account_id(self):
        wb = load_workbook(self.fin)
        return wb['Lista Movimenti']['D8'].value

    def _get_currency(self):
        trans_map = {'Euro': 'EUR'}
        wb = load_workbook(self.fin)
        val = wb['Lista Movimenti']['D22'].value
        return trans_map[val]

    def _get_movimenti(self):
        wb = load_workbook(self.fin)
        starting_column = 'A'
        ending_column = 'G'
        starting_row = 30
        offset = 0

        while True:
            data = wb['Lista Movimenti'][
                        '{}{}:{}{}'.format(starting_column,
                                           starting_row + offset,
                                           ending_column,
                                           starting_row + offset
                                           )]

            values = [*map(lambda x: x.value, data[0])]
            logging.debug(values)
            if not values[0]:
                break
            else:
                yield Movimento(*values)
                offset += 1

    def _get_transaction_type(movimento):
        trans_map = {'Pagamento pos': 'POS',
                     'Pagamento effettuato su pos estero': 'POS',
                     'Accredito beu con contabile': 'XFER',
                     'Canone mensile base e servizi aggiuntivi': 'SRVCHG',
                     'Prelievo carta debito su banche del gruppo': 'CASH',
                     'Prelievo carta debito su banche italia/sepa': 'CASH',
                     'Versamento contanti su sportello automatico':
                     'DIRECTDEP',
                     'Comm.prelievo carta debito italia/sepa': 'SRVCHG',
                     'Commiss. su beu internet banking': 'SRVCHG',
                     'Pagamento telefono': 'PAYMENT',
                     'Pagamento mav via internet banking': 'PAYMENT',
                     'Pagamento bolletta cbill': 'PAYMENT',
                     'Beu tramite internet banking': 'PAYMENT',
                     'Commissione bolletta cbill': 'SRVCHG',
                     'Storno pagamento pos': 'POS',
                     'Versamento contanti su sportello automatico': 'ATM'}
        return trans_map[movimento.descrizione]

    def _get_start_balance(self):
        wb = load_workbook(self.fin)
        return float(wb['Lista Movimenti']['E11'].value)

    def _get_end_balance(self):
        wb = load_workbook(self.fin)
        return float(wb['Lista Movimenti']['E12'].value)

    def _get_start_date(self):
        wb = load_workbook(self.fin)
        date = wb['Lista Movimenti']['D11'].value
        return datetime.strptime(date, '%d.%m.%Y')

    def _get_end_date(self):
        wb = load_workbook(self.fin)
        date = wb['Lista Movimenti']['D12'].value
        return datetime.strptime(date, '%d.%m.%Y')
