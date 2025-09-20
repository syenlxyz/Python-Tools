from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from send2trash import send2trash
from win32com.client import Dispatch
import re
import psutil
import shutil

WdFindWrap  = {
    'wdFindAsk': 2,
    'wdFindContinue': 1,
    'wdFindStop': 0
}

WdReplace = {
    'wdReplaceAll': 2,
    'wdReplaceNone': 0,
    'wdReplaceOne': 1
}

WdHeaderFooterIndex = {
    'wdHeaderFooterEvenPages': 3,
    'wdHeaderFooterFirstPage': 2,
    'wdHeaderFooterPrimary': 1,
}

WdSaveFormat = {
    'wdFormatDocument': 0,
    'wdFormatHTML': 8,
    'wdFormatDocumentDefault': 16,
    'wdFormatPDF': 17
}

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        shutil.copytree(str(input_path), str(output_path))
    else:
        send2trash(output_path)
        shutil.copytree(str(input_path), str(output_path))
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = list(output_path.glob('**/*.docx'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text(f'Processing: done'),
        **options
    )
    
    data = {
        'COMPANY': 'SIN KEE CHIANG ESTATE SDN BHD',
        'NUMBER': '257853-P',
        'PREPARE': 'Poh Moi Hua',
        'APPROVE': 'Wong Gin Poh',
        'DATE': '06/01/2025'
    }
    
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    for file_path in results:
        results.text(f'Processing: {file_path.name}')
        try:
            wrd.Documents.Open(str(file_path))
        except:
            process_list = list(psutil.process_iter())
            for process in process_list:
                if process.name() == 'WinWord.exe':
                    process.terminate()
            wrd.Documents.Open(str(file_path))
        
        keys = list(data.keys())
        for key in keys:
            old = f'${{key}}'
            new = data[key]
            params = {
                'FindText': old,
                'MatchCase': False,
                'MatchWholeWord': True,
                'MatchWildcards': False,
                'MatchSoundsLike': False,
                'MatchAllWordForms': False,
                'Forward': True,
                'Wrap': WdFindWrap['wdFindContinue'],
                'Format': True,
                'ReplaceWith': new,
                'Replace': WdReplace['wdReplaceAll']
            }
            wrd.Selection.Find.Execute(**params)
            wrd.ActiveDocument.Sections(1).Headers(WdHeaderFooterIndex['wdHeaderFooterPrimary']).Range.Find.Execute(**params)
        wrd.ActiveDocument.Close(SaveChanges=True)
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