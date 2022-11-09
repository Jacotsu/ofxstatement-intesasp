#!/usr/bin/env python3.6

import logging
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal

from ofxstatement.parser import StatementParser
from ofxstatement.plugin import Plugin
from ofxstatement.statement import (Statement, StatementLine,
                                    generate_transaction_id)
from openpyxl import load_workbook

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger('IntesaSP')


@dataclass
class Movimento:
    data_contabile: datetime
    data_valuta: datetime
    descrizione: str
    accrediti: Decimal
    addebiti: Decimal
    descrizione_estesa: str
    mezzo: str

    def __post_init__(self):
        # Modificare descrizione_estesa per comprendere sempre la descrizione
        self.descrizione_estesa = "({}) {}".format(self.descrizione, self.descrizione_estesa)


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
                                  Decimal(mov.accrediti) if mov.accrediti else
                                  Decimal(mov.addebiti))
        stat_line.id = generate_transaction_id(stat_line)
        stat_line.date_user = mov.data_valuta
        stat_line.trntype = IntesaSanPaoloXlsxParser._get_transaction_type(mov)
        logging.debug(stat_line)
        return stat_line

    def _get_account_id(self):
        wb = load_workbook(self.fin)
        return wb['Lista Movimenti']['D8'].value

    def _get_currency(self):
        # !!! Write All key value lower case !!!
        trans_map = {'euro': 'EUR',
                     'eur': 'EUR'}
        wb = load_workbook(self.fin)
        val = wb['Lista Movimenti']['D22'].value
        return trans_map[val.lower()]

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
        # OFX Spec https://financialdataexchange.org/ofx
        # 11.4.4.3 Transaction Types Used in <TRNTYPE>
        # !!! Write All key value lower case !!!

        trans_map = {
            # Pagamenti POS
            'pagamento pos': 'POS',
            'pagamento tramite pos': 'POS',
            'pagamento effettuato su pos estero': 'POS',
            'storno pagamento pos': 'POS',
            'storno pagamento pos estero': 'POS',
            # Accrediti
            'canone mensile base e servizi aggiuntivi': 'SRVCHG',
            'prelievo carta debito su banche del gruppo': 'CASH',
            'prelievo carta debito su banche italia/sepa': 'CASH',
            'comm.prelievo carta debito italia/sepa': 'SRVCHG',
            'stipendio o pensione': 'CREDIT',
            # Ricariche Cellulari
            'pagamento mav via internet banking': 'PAYMENT',
            'pagamento bolletta cbill': 'PAYMENT',
            'pagamento telefono': 'PAYMENT',
            'ricarica tramite internet:vodafonecard': 'PAYMENT',
            'ricarica tramite internet:windtre': 'PAYMENT',
            # Commissioni Pagamento
            'commissione bolletta cbill': 'FEE',
            'commissioni bollettino postale via internet': 'FEE',
            'commissioni e spese adue': 'FEE',
            'commissioni su pagamento via internet': 'FEE',
            'imposta di bollo e/c e rendiconto': 'FEE',
            'commiss. su beu internet banking': 'FEE',
            # Operazioni Bancarie
            'accredito beu con contabile': 'XFER',
            'beu tramite internet banking': 'XFER',
            'versamento contanti su sportello automatico': 'ATM',
            'canone annuo o-key sms': 'SRVCHG',
            'pagamento adue': 'DIRECTDEBIT',
            'rata bonif. periodico con contab.': 'REPEATPMT',
            'bonifico in euro verso ue/sepa canale telem.': 'PAYMENT',
            'accredito bonifico istantaneo': 'DIRECTDEP',
            'pagamenti disposti su circuito fast pay': 'PAYMENT',
            'pagamento bollettino postale via internet': 'PAYMENT',
            'pagamento via internet': 'DIRECTDEBIT',
            # Other
            'donazione preautorizzata ad ente no profit': 'DIRECTDEBIT',
            'add. deleghe fisco/inps/regioni': 'DEBIT',
            'pagamento delega f24 via internet banking': 'PAYMENT',
        }
        currentTransition = trans_map.get(movimento.descrizione.lower())
        if currentTransition is None:
          currentTransition = 'DIRECTDEBIT'
          print("Warning!! The transition type '{}' is not present yet on code!!\n" \
                "PLESE report this issue on GitHub Repository 'https://github.com/Jacotsu/ofxstatement-intesasp/issues' to Help US" \
                "Now for that Transition will be assign the default type: {}".format(movimento.descrizione, currentTransition))
        return currentTransition

    def _get_start_balance(self):
        wb = load_workbook(self.fin)
        return Decimal(wb['Lista Movimenti']['E11'].value)

    def _get_end_balance(self):
        wb = load_workbook(self.fin)
        return Decimal(wb['Lista Movimenti']['E12'].value)

    def _get_start_date(self):
        wb = load_workbook(self.fin)
        date = wb['Lista Movimenti']['D11'].value
        return datetime.strptime(date, '%d.%m.%Y')

    def _get_end_date(self):
        wb = load_workbook(self.fin)
        date = wb['Lista Movimenti']['D12'].value
        return datetime.strptime(date, '%d.%m.%Y')
