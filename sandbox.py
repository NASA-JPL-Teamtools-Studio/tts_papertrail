import pdb
import pytest
from pathlib import Path

from data_utils.core.evr import EvrContainer
from data_utils.core.eha import EhaContainer
from dexter.src.dexter import Dexter
from dexter.dispositioners.evrs import BulkEvrDispositioner
from dexter.dispositioners.eha import LadEhaDispositioner
from dexter.src.dispo import DISPO_FORMAT #TO DO: Clean this up... this is vestigal because it became data_utils
from papertrail.excelerate import ExcelReport
from papertrail.logger import create_logger

logger = create_logger(name='papertrail.test_excelerate')

TEST_FILE_DIR = Path(__file__).parent.joinpath('test_files/excelerate')

class TestBulkEvrDispositioner(BulkEvrDispositioner):
    CSV_FILEPATH = TEST_FILE_DIR.joinpath('evr_autodispositions.csv')

class LadBulkEhaDispositioner(LadEhaDispositioner):
    CSV_FILEPATH = TEST_FILE_DIR.joinpath('eha_autodispositions.csv')

class EvrBulkDispoTestDexter(Dexter):
    def __init__(self, evrs=None, eha=None):
        super().__init__()        
        self.init_data(EvrContainer, evrs)
        self.init_data(EhaContainer, eha)
        self.init_dispositioner(TestBulkEvrDispositioner)
        self.init_dispositioner(LadBulkEhaDispositioner)

evr_container = EvrContainer(
    csv_path=TEST_FILE_DIR.joinpath('evrs.csv'), 
    cast_fields=True
    )    

eha_container = EhaContainer(
    csv_path=TEST_FILE_DIR.joinpath('dn_chanvals_no_tolerance.csv'), 
    cast_fields=True
    )    

dex = EvrBulkDispoTestDexter(evr_container, eha_container)
dex.dispo_format = DISPO_FORMAT.EXCEL
dex.disposition_all()
dex.stamp_all()


report = ExcelReport(name='Unit Test Report', author='pytest')
report.add_entry('EVR Autodispositions', evr_container)
report.add_entry('EHA Autodispositions', eha_container)
report.commit()








