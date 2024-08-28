from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

iPageOption = {
    'PDBeforeFirstPage': -1,
    'PDLastPage': -2,
    'PDAllPages': -3,
    'PDOddPagesOnly': -4,
    'PDEvenPagesOnly': -5
}

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = list(input_path.glob('*.pdf'))
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing PDF Document: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Printing PDF Document: {file_path.name}')
        print_pdf(file_path)

def print_pdf(file_path):
    app = Dispatch('AcroExch.App')
    app.Hide()
    
    avDoc = Dispatch('AcroExch.AVDoc')
    avDoc.Open(file_path.as_posix(), '')
    
    pdDoc = avDoc.GetPDDoc()
    num_page = pdDoc.GetNumPages()
    
    params = {
        'nFirstPage': 0,
        'nLastPage': num_page - 1,
        'nPSLevel': 3,
        'bBinaryOk': 0,
        'bShrinkToFit': 0,
        'bReverse': False,
        'bFarEastFontOpt': 0,
        'bEmitHalftones': 0,
        'iPageOption': iPageOption['PDAllPages']
    }
    
    avDoc.PrintPagesEx(**params)
    avDoc.Close(True)
    app.MenuItemExecute('Quit')

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')