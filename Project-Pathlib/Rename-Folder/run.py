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
    
    folder_list = []
    path_list = list(output_path.glob('**/*'))
    for path in path_list:
        if path.is_dir():
            folder_list.append(path)
    
    results = alive_it(
        folder_list, 
        len(folder_list), 
        finalize=lambda bar: bar.text(f'Processing: done'),
        **options
    )
    
    old = '5C'
    new = '5C0'
    for folder_path in results:
        results.text(f'Processing: {folder_path.name}')
        target_path = folder_path.parent / folder_path.name.replace(old, new)
        folder_path.rename(target_path)

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')