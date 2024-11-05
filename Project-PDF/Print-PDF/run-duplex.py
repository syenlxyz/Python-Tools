from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch
from win32print import GetDefaultPrinter, OpenPrinter, GetPrinter
import psutil

iPageOption = {
    'PDAllPages': -3,
    'PDOddPagesOnly': -4,
    'PDEvenPagesOnly': -5
}

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = list(input_path.glob('**/*.pdf'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing PDF: done'),
        **options
    )
    
    process_list = list(psutil.process_iter())
    for process in process_list:
        if process.name() == 'Acrobat.exe':
            process.terminate()
    
    params = {
        'nFirstPage': 0,
        'nLastPage': None,
        'nPSLevel': 3,
        'bBinaryOk': False,
        'bShrinkToFit': True,
        'bReverse': False,
        'bFarEastFontOpt': False,
        'bEmitHalftones': False,
        'iPageOption': iPageOption['PDAllPages']
    }
    
    app = Dispatch('AcroExch.App')
    app.Hide()
    avDoc = Dispatch('AcroExch.AVDoc')
    for file_path in results:
        results.text(f'Printing PDF: {file_path.name}')
        avDoc.Open(file_path, file_path)
        pdDoc = avDoc.GetPDDoc()
        num_page = pdDoc.GetNumPages()
        params['nLastPage'] = num_page - 1
        avDoc.PrintPagesEx(**params)
        avDoc.Close(bNoSave=True)
    
    process_list = list(psutil.process_iter())
    for process in process_list:
        if process.name() == 'Acrobat.exe':
            process.terminate()

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')