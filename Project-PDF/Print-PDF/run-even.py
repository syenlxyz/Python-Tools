from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter
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
        finalize=lambda bar: bar.text('Printing PDF (Even Pages Only): done'),
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
        'bReverse': True,
        'bFarEastFontOpt': False,
        'bEmitHalftones': False,
        'iPageOption': iPageOption['PDEvenPagesOnly']
    }
    
    default_printer = GetDefaultPrinter()
    handle = OpenPrinter(default_printer)
    level = 2
    printer_info = GetPrinter(handle, level)
    printer_info['pDevMode'].Duplex = 1
    
    app = Dispatch('AcroExch.App')
    app.Hide()
    avDoc = Dispatch('AcroExch.AVDoc')
    for file_path in results:
        results.text(f'Printing PDF Document (Even Pages Only): {file_path.name}')
        reader = PdfReader(file_path)
        num_page = reader.get_num_pages()
        if num_page % 2:
            writer = PdfWriter(reader)
            writer.add_blank_page()
            temp_path = input_path / 'temp.pdf'
            with open(temp_path, 'wb'):
                writer.write(temp_path)
            avDoc.Open(temp_path, temp_path)
            params['nLastPage'] = num_page
            avDoc.PrintPagesEx(**params)
            avDoc.Close(bNoSave=True)
            temp_path.unlink()
        else:
            avDoc.Open(file_path, file_path)
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