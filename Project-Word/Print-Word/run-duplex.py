from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch
from win32print import GetDefaultPrinter, OpenPrinter, GetPrinter

WdPrintOutPages = {
    'AllPages': 0,
    'EvenPagesOnly': 2,
    'OddPagesOnly': 1
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
    
    file_list = []
    path_list = list(input_path.glob('**/*'))
    for path in path_list:
        if path.suffix in ['.doc', '.docx']:
            file_list.append(path)
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing Word: done'),
        **options
    )
    
    params = {
        'Background': False,
        'Append': False,
        'Range': 0,
        'OutputFileName': '',
        'From': '',
        'To': '',
        'Item': 0,
        'Copies': 1,
        'Pages': '',
        'PageType': WdPrintOutPages['AllPages'],
        'PrintToFile': False,
        'Collate': True
    }
    
    default_printer = GetDefaultPrinter()
    handle = OpenPrinter(default_printer)
    level = 2
    printer_info = GetPrinter(handle, level)
    default_duplex = printer_info['pDevMode'].Duplex
    
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    wrd.Options.PrintReverse = False
    for file_path in results:
        results.text(f'Printing Word: {file_path.name}')
        wrd.Documents.Open(str(file_path))
        wrd.ActiveDocument.PrintOut(**params)
        if wrd.ActiveDocument.PageSetup.Orientation:
            printer_info['pDevMode'].Duplex = 3
        else:
            printer_info['pDevMode'].Duplex = 2
        wrd.ActiveDocument.Close(SaveChanges=False)
    wrd.Quit()
    printer_info['pDevMode'].Duplex = default_duplex

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')