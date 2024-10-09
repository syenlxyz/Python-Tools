from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from win32com.client import Dispatch
from win32print import SetDefaultPrinter
import psutil

PageOption = {
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
    
    target_printer = 'EPSON L3150 Series'
    SetDefaultPrinter(target_printer)
    
    process_list = list(psutil.process_iter())
    for process in process_list:
        if process.name() == 'Acrobat.exe':
            process.terminate()
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing PDF Document (Odd Pages Only): done'),
        **options
    )
    for file_path in results:
        results.text(f'Printing PDF Document (Odd Pages Only): {file_path.name}')
        print_pdf(file_path, PageOption['PDOddPagesOnly'])

    input('Press ENTER to proceed...')
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing PDF Document (Even Pages Only): done'),
        **options
    )
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
            
            print_pdf(temp_path, PageOption['PDEvenPagesOnly'])
            temp_path.unlink()
        else:
            print_pdf(file_path, PageOption['PDEvenPagesOnly'])
    

def print_pdf(file_path, iPageOption):
    app = Dispatch('AcroExch.App')
    app.Hide()
    
    avDoc = Dispatch('AcroExch.AVDoc')
    avDoc.Open(file_path, file_path)
    
    pdDoc = avDoc.GetPDDoc()
    num_page = pdDoc.GetNumPages()
    
    params = {
        'nFirstPage': 0,
        'nLastPage': num_page - 1,
        'nPSLevel': 3,
        'bBinaryOk': False,
        'bShrinkToFit': True,
        'bReverse': True,
        'bFarEastFontOpt': False,
        'bEmitHalftones': False,
        'iPageOption': iPageOption
    }
    avDoc.PrintPagesEx(**params)
    avDoc.Close(True)
    
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