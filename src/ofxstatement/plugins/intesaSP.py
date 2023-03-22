#!/usr/bin/env python3.6

import logging
import os
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from collections.abc import Iterator

from ofxstatement.parser import StatementParser
from ofxstatement.plugin import Plugin
from ofxstatement.statement import (Statement, StatementLine,
                                    generate_transaction_id)

from openpyxl import load_workbook

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger('IntesaSP')

"""
Data wrapper general
"""


class Movimento:
    stat_line: StatementLine = None


"""
Data wrapper for Version 1 (Movimenti_Conto_<DATE>.xlsx)
"""


@dataclass
class Movimento_V1(Movimento):
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

        # Una volta raccolti i dati, li formatto nello standard corretto
        self.stat_line = StatementLine(None,
                                       self.data_contabile,
                                       self.descrizione_estesa,
                                       Decimal(self.accrediti) if self.accrediti else Decimal(self.addebiti))
        self.stat_line.id = generate_transaction_id(self.stat_line)
        self.stat_line.date_user = self.data_valuta
        self.stat_line.trntype = self._get_transaction_type()

    def _get_transaction_type(self):
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
        currentTransition = trans_map.get(self.descrizione.lower())
        if currentTransition is None:
            currentTransition = 'DIRECTDEBIT'
            print("Warning!! The transition type '{}' is not present yet on code!!\n" \
                  "PLESE report this issue on GitHub Repository 'https://github.com/Jacotsu/ofxstatement-intesasp/issues' to Help US" \
                  "Now for that Transition will be assign the default type: {}".format(self.descrizione,
                                                                                       currentTransition))
        return currentTransition


class IntesaSanPaoloPlugin(Plugin):

    def get_parser(self, filename):
        parser = IntesaSanPaoloXlsxParser(filename)
        parser.statement.bank_id = self.settings.get('BIC', 'IntesaSP')
        return parser


class IntesaSanPaoloXlsxParser(StatementParser):
    excelVersion: int = None
    wb = None

    def __init__(self, filename):
        self.fin = filename
        self.wb = load_workbook(self.fin)
        if 'Lista Movimenti' in self.wb.sheetnames:
            print('"Lista Movimenti" exists, excel version 1 parse will be used')
            self.excelVersion = 1
        else:
            print('No know sheet found, impossible continue')
            exit(os.EX_IOERR)

        self.statement = Statement()
        self.statement.account_id = self._get_account_id()
        self.statement.currency = self._get_currency()
        self.statement.start_balance = self._get_start_balance()
        self.statement.start_date = self._get_start_date()
        self.statement.end_balance = self._get_end_balance()
        self.statement.end_date = self._get_end_date()
        logging.debug(self.statement)

    """
    Override method, use to obrain iterable object consisting of a line per transaction
    """

    def split_records(self):
        if self.excelVersion == 1:
            return self._get_movimenti_V1()
        else:
            return None

    # TODO: Capire se serve
    """
    Override method, use to generate Statement object
    """

    def parse(self):
        return super(IntesaSanPaoloXlsxParser, self).parse()

    """
    Override use to Parse given transaction line and return StatementLine object
    """

    def parse_record(self, mov: Movimento):
        logging.debug(mov.stat_line)
        return mov.stat_line

    def _get_account_id(self):
        if self.excelVersion == 1:
            return self.wb['Lista Movimenti']['D8'].value
        else:
            return None

    def _get_currency(self):
        # !!! Write All key value lower case !!!
        trans_map = {'euro': 'EUR',
                     'eur': 'EUR'}
        if self.excelVersion == 1:
            val = self.wb['Lista Movimenti']['D22'].value
            return trans_map[val.lower()]
        else:
            return None

    def _get_movimenti_V1(self) -> Iterator[Movimento_V1]:
        starting_column = 'A'
        ending_column = 'G'
        starting_row = 30
        offset = 0

        while True:
            data = self.wb['Lista Movimenti'][
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
                yield Movimento_V1(*values)
                offset += 1

    def _get_start_balance(self) -> Decimal:
        if self.excelVersion == 1:
            return Decimal(self.wb['Lista Movimenti']['E11'].value)
        else:
            return None

    def _get_end_balance(self) -> Decimal:
        if self.excelVersion == 1:
            return Decimal(self.wb['Lista Movimenti']['E12'].value)
        else:
            return None

    def _get_start_date(self) -> datetime:
        if self.excelVersion == 1:
            date = self.wb['Lista Movimenti']['D11'].value
            return datetime.strptime(date, '%d.%m.%Y')
        else:
            return None

    def _get_end_date(self) -> datetime:
        if self.excelVersion == 1:
            date = self.wb['Lista Movimenti']['D12'].value
            return datetime.strptime(date, '%d.%m.%Y')
        else:
            return None
