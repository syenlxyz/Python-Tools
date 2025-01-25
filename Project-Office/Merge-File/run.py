## Unresolved issue: files with same name
from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from send2trash import send2trash
import shutil

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        output_path.mkdir()
    else:
        send2trash(output_path)
        output_path.mkdir()
    
    file_list = []
    path_list = list(input_path.glob('**/*'))
    for path in path_list:
        if path.is_file():
            file_list.append(path)
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Merging File: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Merging File: {file_path.name}')
        target_path = output_path / file_path.name
        shutil.copy2(str(file_path), str(target_path))

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')