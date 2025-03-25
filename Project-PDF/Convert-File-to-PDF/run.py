from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from send2trash import send2trash
from win32com.client import Dispatch
import psutil
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
    
    file_list = []
    path_list = list(output_path.glob('**/*'))
    suffix_list = ['.html', '.htm', '.jpg', '.jpeg', '.xls', '.xlsx', '.ppt', '.pptx', '.doc', '.docx', '.png', '.txt']
    for path in path_list:
        if path.suffix in suffix_list:
            file_list.append(path)
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Converting File to PDF: done'),
        **options
    )
    
    process_list = list(psutil.process_iter())
    for process in process_list:
        if process.name() == 'Acrobat.exe':
            process.terminate()
    
    app = Dispatch('AcroExch.App')
    app.Hide()
    avDoc = Dispatch('AcroExch.AVDoc')
    for file_path in results:
        results.text(f'Converting File to PDF: {file_path.name}')
        avDoc.Open(file_path, file_path)
        pdDoc = avDoc.GetPDDoc()
        target_path = file_path.with_suffix('.pdf')
        pdDoc.Save(1, target_path)
        avDoc.Close(bNoSave=True)
    
    process_list = list(psutil.process_iter())
    for process in process_list:
        if process.name() == 'Acrobat.exe':
            process.terminate()
    
    for file_path in file_list:
        send2trash(file_path)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')