from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

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
    
    file_list = list(input_path.glob('*.doc')) + list(input_path.glob('*.docx'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Converting Word to PDF: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Converting Word to PDF: {file_path.name}')
        word_to_pdf(file_path, output_path)

def word_to_pdf(file_path, output_path):
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    pdf_path = output_path / file_path.with_suffix('.pdf').name
    doc = wrd.Documents.Open(str(file_path))
    doc.SaveAs(str(pdf_path), FileFormat=17)
    wrd.Quit(SaveChanges=False)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')