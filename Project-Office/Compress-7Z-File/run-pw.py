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
    
    folder_list = list(input_path.iterdir())
    results = alive_it(
        folder_list, 
        len(folder_list), 
        finalize=lambda bar: bar.text('Compressing 7Z File with Password: done'),
        **options
    )
    
    for folder_path in results:
        results.text(f'Compressing 7Z File with Password: {folder_path.name}')
        target_path = output_path / folder_path.with_suffix('.7z').name
        password = folder_path.name
        with SevenZipFile(target_path, 'w', password=password) as file:
            file.writeall(folder_path)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')