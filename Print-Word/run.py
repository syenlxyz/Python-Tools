from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

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
    
    file_list = list(input_path.glob('*.doc')) + list(input_path.glob('*.docx'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing Word: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Printing Word: {file_path.name}')
        print_word(file_path)

def print_word(file_path):
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    wrd.Options.PrintReverse = False
    
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
    
    wrd.Documents.Open(str(file_path))
    wrd.PrintOut(**params)
    wrd.Quit(SaveChanges=False)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')