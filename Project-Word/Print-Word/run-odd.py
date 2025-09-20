## Unresolved issue: run-odd-rev is not created
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
        finalize=lambda bar: bar.text('Processing: done'),
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
        'PageType': WdPrintOutPages['OddPagesOnly'],
        'PrintToFile': False,
        'Collate': True
    }
    
    default_printer = GetDefaultPrinter()
    handle = OpenPrinter(default_printer)
    printer_info = GetPrinter(handle, 2)
    printer_info['pDevMode'].Duplex = 1
    
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    wrd.Options.PrintReverse = True
    for file_path in results:
        results.text(f'Processing: {file_path.name}')
        wrd.Documents.Open(str(file_path))
        wrd.ActiveDocument.PrintOut(**params)
        wrd.ActiveDocument.Close(SaveChanges=False)
    wrd.Options.PrintReverse = False
    wrd.Quit()

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')