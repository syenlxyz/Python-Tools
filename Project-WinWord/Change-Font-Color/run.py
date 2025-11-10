from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from send2trash import send2trash
from win32com.client import Dispatch
import shutil

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
    
    FONT_COLOR = 0x000000
    
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    for file_path in results:
        results.text(f'Processing: {file_path.name}')
        doc = wrd.Documents.Open(str(file_path))
        
        doc.Content.Select()
        doc.Content.Font.Color = FONT_COLOR
        
        doc.Sections(1).Headers(1).Range.Select()
        doc.Sections(1).Headers(1).Range.Font.Color = FONT_COLOR
        
        doc.Sections(1).Footers(1).Range.Select()
        doc.Sections(1).Footers(1).Range.Font.Color = FONT_COLOR
        
        doc.Close(SaveChanges=True)
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