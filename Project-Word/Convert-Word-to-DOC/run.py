from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
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
        shutil.rmtree(str(output_path))
        shutil.copytree(str(input_path), str(output_path))
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = list(input_path.glob('**/*.docx'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Converting Word to DOC: done'),
        **options
    )
    
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    for file_path in results:
        results.text(f'Converting Word to DOC: {file_path.name}')
        wrd.Documents.Open(str(file_path))
        target_path = file_path.with_suffix('.doc')
        wrd.ActiveDocument.SaveAs2(str(target_path), FileFormat=0)
        wrd.ActiveDocument.Close(SaveChanges=False)
        file_path.unlink()
    wrd.Quit()

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')