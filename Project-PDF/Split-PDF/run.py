from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from send2trash import send2trash
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
    
    file_list = list(output_path.glob('**/*.pdf'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Processing: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Processing: {file_path.name}')
        folder_path = file_path.with_suffix('')
        folder_path.mkdir()
        reader = PdfReader(file_path)
        pages = list(reader.pages)
        for index, page in enumerate(pages):
            writer = PdfWriter()
            writer.add_page(page)
            target_path = folder_path / f'page{index + 1}.pdf'
            with open(target_path, 'wb') as file:
                writer.write(file)
        send2trash(file_path)

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')