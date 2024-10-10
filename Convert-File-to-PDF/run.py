from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch
import psutil

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        output_path.mkdir()
    else:
        file_list = list(output_path.iterdir())
        for file_path in file_list:
            file_path.unlink()
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = []
    ext_list = ['html', 'htm', 'jpg', 'jpeg', 'xls', 'xlsx', 'ppt', 'pptx', 'doc', 'docx', 'png', 'txt']
    for ext in ext_list:
        item = list(input_path.glob(f'*.{ext}'))
        file_list.extend(item)
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Converting File to PDF: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Converting File to PDF: {file_path.name}')
        file_to_pdf(file_path, output_path)

def file_to_pdf(file_path, output_path):
    app = Dispatch('AcroExch.App')
    app.Hide()
    
    avDoc = Dispatch('AcroExch.AVDoc')
    avDoc.Open(file_path, file_path)
    pdDoc = avDoc.GetPDDoc()
    
    pdf_path = output_path / file_path.with_suffix('.pdf').name
    pdDoc.Save(1, pdf_path)
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