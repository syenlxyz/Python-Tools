from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

WdPrintOutRange = {
    'AllDocument': 0,
    'CurrentPage': 2,
    'FromTo': 3,
    'RangeOfPages': 4,
    'Selection': 1
}

WdPrintOutItem = {
    'AutoTextEntries': 4,
    'Comments': 2,
    'DocumentContent': 0,
    'DocumentWithMarkup': 7,
    'Envelope': 6,
    'KeyAssignments': 5,
    'Markup': 2,
    'Properties': 1,
    'Styles': 3
}

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
    wrd.Options.PrintReverse = True
    
    params = {
        'Background': False,
        'Append': False,
        'Range': WdPrintOutRange['AllDocument'],
        'OutputFileName': '',
        'From': '',
        'To': '',
        'Item': WdPrintOutItem['DocumentContent'],
        'Copies': 1,
        'Pages': '',
        'PageType': WdPrintOutPages['AllPages'],
        'PrintToFile': False,
        'Collate': True
    }
    
    wrd.Documents.Open(file_path.as_posix(), ReadOnly=True)
    wrd.PrintOut(**params)
    wrd.Quit(SaveChanges=False)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')