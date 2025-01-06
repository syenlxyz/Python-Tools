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
    
    folder_list = []
    path_list = list(input_path.iterdir())
    for path in path_list:
        if path.is_dir():
            folder_list.append(path)
    
    for folder_path in folder_list:
        file_list = list(folder_path.glob('**/*.7z'))
        results = alive_it(
            file_list, 
            len(file_list), 
            finalize=lambda bar: bar.text(f'Extracting 7Z File with Password for {folder_path.name}: done'),
            **options
        )
        
        for file_path in results:
            results.text(f'Extracting 7Z File with Password for {folder_path.name}: {file_path.name}')
            target_path = output_path / file_path.stem
            password = folder_path.name
            with SevenZipFile(file_path, 'r', password=password) as file:
                file.extractall(target_path)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')