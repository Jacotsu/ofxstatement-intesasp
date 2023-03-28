#!/usr/bin/env python3.6

import logging
import os
from collections.abc import Iterator
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

GITHUB_URL = "https://github.com/Jacotsu/ofxstatement-intesasp/issues"

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
        # Map value present in ~/.local/lib/python3.10/site-packages/ofxstatement/statement.py

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
            print(f"Warning!! The transition type '{self.descrizione}' is not present yet on code!!\n" \
                  f"PLESE report this issue on GitHub Repository '{GITHUB_URL}' to Help US" \
                  f"Now for that Transition will be assign the default type: {currentTransition}")
        return currentTransition


@dataclass
class Movimento_V2(Movimento):
    data: datetime
    operazione: str
    dettagli: str
    conto_carta: str
    contabilizzazione: str
    categoria: str
    valuta: str
    importo: Decimal

    def __post_init__(self):
        # Modificare descrizione_estesa per comprendere sempre la descrizione
        descrizione_estesa = f"[({self.categoria})-({self.operazione})] {self.dettagli}"

        # Una volta raccolti i dati, li formatto nello standard corretto
        self.stat_line = StatementLine(None, self.data, descrizione_estesa, Decimal(self.importo))
        self.stat_line.id = generate_transaction_id(self.stat_line)
        self.stat_line.date_user = self.data
        self.stat_line.trntype = self._get_transaction_type()

    def _get_transaction_type(self):
        # OFX Spec https://financialdataexchange.org/ofx
        # 11.4.4.3 Transaction Types Used in <TRNTYPE>
        # !!! Write All key value lower case !!!
        # Map value present in ~/.local/lib/python3.10/site-packages/ofxstatement/statement.py

        categoryMap = {
            'addebiti vari': 'DEBIT',
            'abbigliamento e accessori': 'POS',
            'altre uscite': 'POS',
            'associazioni': 'DIRECTDEBIT',
            'bonifici in uscita': 'XFER',
            'bonifici ricevuti': 'XFER',
            'carburanti': 'PAYMENT',
            'casa varie': 'POS',
            'cellulare': 'SRVCHG',
            'cliniche': 'POS',
            'corsi e sport': 'POS',
            'cura della persona': 'POS',
            'domiciliazioni e utenze': 'DIRECTDEBIT',
            'donazioni': 'DIRECTDEBIT',
            'farmacia': 'POS',
            'generi alimentari e supermercato': 'POS',
            'hi-tech e informatica': 'POS',
            'imposte sul reddito e tasse varie': 'FEE',
            'imposte, bolli e commissioni': 'FEE',
            'pedaggi e telepass': 'FEE',
            'rate mutuo e finanziamento': 'DIRECTDEBIT',
            'regali ricevuti': 'CREDIT',
            'rimborsi spese e storni': 'CREDIT',
            'ristoranti e bar': 'POS',
            'spese mediche': 'POS',
            'spettacoli e musei': 'POS',
            'stipendi e pensioni': 'DIRECTDEP',
            'tv, internet, telefono': 'POS',
            'tabaccai e simili': 'POS',
            'tempo libero varie': 'POS',
            'viaggi e vacanze': 'POS'
        }

        currentTransition = categoryMap.get(self.categoria.lower())
        if currentTransition is None:
            currentTransition = 'CREDIT' if self.importo >= 0 else 'DEBIT'
            print(f"Op: '{self.categoria}' is not present yet, assign '{currentTransition}'")
            print(f"PLESE report this issue on GitHub Repository '{GITHUB_URL}' to Help US")
            print("-" * 60)
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
        if 'Lista Operazione' in self.wb.sheetnames:
            print('"Lista Operazione" exists, excel version 2 parse will be used')
            self.excelVersion = 2
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
        if self.excelVersion == 2:
            return self._get_movimenti_V2()
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

    def parse_record(self, mov: Movimento) -> StatementLine:
        logging.debug(mov.stat_line)
        return mov.stat_line

    def _get_account_id(self) -> str:
        if self.excelVersion == 1:
            return self.wb['Lista Movimenti']['D8'].value
        if self.excelVersion == 2:
            return self.wb['Lista Operazione']['C7'].value.split(" ")[1]
        else:
            return None

    def _get_currency(self) -> str:
        # !!! Write All key value lower case !!!
        trans_map = {'euro': 'EUR',
                     'eur': 'EUR'}
        if self.excelVersion == 1:
            val = self.wb['Lista Movimenti']['D22'].value
        if self.excelVersion == 2:
            # Suppose all operation have same currency
            val = self.wb['Lista Operazione']['G20'].value  # First operation
        else:
            return None
        return trans_map[val.lower()]

    def _get_start_balance(self) -> Decimal:
        if self.excelVersion == 1:
            return Decimal(self.wb['Lista Movimenti']['E11'].value)
        if self.excelVersion == 2:
            return None
        else:
            return None

    def _get_end_balance(self) -> Decimal:
        if self.excelVersion == 1:
            return Decimal(self.wb['Lista Movimenti']['E12'].value)
        if self.excelVersion == 2:
            return None
        else:
            return None

    def _get_start_date(self) -> datetime:
        if self.excelVersion == 1:
            date = self.wb['Lista Movimenti']['D11'].value
            formatData = '%d.%m.%Y'
        if self.excelVersion == 2:
            date = self.wb['Lista Operazione']['C16'].value
            formatData = '%d/%m/%Y'
        else:
            return None
        return datetime.strptime(date, formatData)

    def _get_end_date(self) -> datetime:
        if self.excelVersion == 1:
            date = self.wb['Lista Movimenti']['D12'].value
            formatData = '%d.%m.%Y'
        if self.excelVersion == 2:
            date = self.wb['Lista Operazione']['C17'].value
            formatData = '%d/%m/%Y'
        else:
            return None
        return datetime.strptime(date, formatData)

    """ Private method to parse all record lines """

    def _get_movimenti_V1(self) -> Iterator[Movimento_V1]:
        starting_column = 'A'
        ending_column = 'G'
        starting_row = 30
        offset = 0

        while True:
            data = self.wb['Lista Movimenti']['{}{}:{}{}'
            .format(starting_column, starting_row + offset, ending_column, starting_row + offset)]
            offset += 1

            values = [*map(lambda x: x.value, data[0])]
            logging.debug(values)
            if not values[0]:
                # Complete transaction table finish, end loop
                break
            else:
                yield Movimento_V1(*values)

    def _get_movimenti_V2(self) -> Iterator[Movimento_V2]:
        starting_column = 'A'
        ending_column = 'H'
        starting_row = 20
        offset = 0

        while True:
            data = self.wb['Lista Operazione']['{}{}:{}{}'
            .format(starting_column, starting_row + offset, ending_column, starting_row + offset)]
            offset += 1
            values = [*map(lambda x: x.value, data[0])]
            logging.debug(values)
            if not values[0]:
                # Table finish, end loop
                break
            if values[4] == "NON CONTABILIZZATO":
                # This tool save only finish transaction
                continue
            else:
                yield Movimento_V2(*values)
