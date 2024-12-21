from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from py7zr import SevenZipFile
import shutil

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        output_path.mkdir()
    else:
        shutil.rmtree(str(output_path))
        output_path.mkdir()
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = list(input_path.glob('**/*.7z'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Extracting 7z Files: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Extracting 7z Files: {file_path.name}')
        target_path = output_path / file_path.stem
        with SevenZipFile(file_path, mode='r') as file:
            if not file.needs_password():
                file.extract(target_path)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')