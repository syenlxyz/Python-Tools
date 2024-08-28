from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

WdStatistic = {
    'Characters': 3,
    'CharactersWithSpaces': 5,
    'FarEastCharacters': 6,
    'Lines': 1,
    'Pages': 2,
    'Paragraphs': 4,
    'Words': 0
}

WdGoToItem = {
    'Bookmark': -1,
    'Comment': 6,
    'Endnote': 5,
    'Equation': 10,
    'Field': 7,
    'Footnote': 4,
    'GrammaticalError': 14,
    'Graphic': 8,
    'Heading': 11,
    'Line': 3,
    'Object': 9,
    'Page': 1,
    'Percent': 12,
    'ProofreadingError': 15,
    'Section': 0,
    'SpellingError': 13,
    'Table': 2
}

WdGoToDirection   = {
    'Absolute': 1,
    'First': 1,
    'Last': -1,
    'Next': 2,
    'Previous': 3,
    'Relative': 2
}

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
    
    file_list = list(input_path.glob('*.doc')) + list(input_path.glob('*.docx'))
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
        finalize=lambda bar: bar.text('Printing Word Document: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Printing Word Document: {file_path.name}')
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
    
    num_page = wrd.ActiveDocument.ComputeStatistics(Statistic=WdStatistic['Pages'])
    if num_page % 2:
        wrd.Selection.GoTo(WdGoToItem['Line'], WdGoToDirection['Last'])
        wrd.Selection.EndKey()
        wrd.Selection.InsertNewPage()
    
    params['PageType'] = WdPrintOutPages['OddPagesOnly']
    wrd.PrintOut(**params)
    input('Press ENTER to continue...')
    params['PageType'] = WdPrintOutPages['EvenPagesOnly']
    wrd.PrintOut(**params)
    
    wrd.Quit(SaveChanges=False)

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')