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
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = []
    folder_list = []
    path_list = list(input_path.iterdir())
    for path in path_list:
        if path.is_file():
            file_list.append(path)
        if path.is_dir():
            folder_list.append(path)
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text(f'Processing {input_path.name}: done'),
        **options
    )
    for file_path in results:
        results.text(f'Processing {input_path.name}: {file_path.name}')
        num_copy = int(file_path.name.split('+')[0])
        prefix = file_path.stem.split('+')[-1]
        for index in range(num_copy):
            target_path = output_path / f'{prefix}_copy{index + 1}{file_path.suffix}'
            shutil.copy2(str(file_path), str(target_path))
        
    for folder_path in folder_list:
        file_list = list(folder_path.iterdir())
        results = alive_it(
            file_list, 
            len(file_list), 
            finalize=lambda bar: bar.text(f'Processing {folder_path.name}: done'),
            **options
        )
        num_copy = int(folder_path.name.split('+')[0])
        for file_path in results:
            results.text(f'Processing {folder_path.name}: {file_path.name}')
            for index in range(num_copy):
                target_path = output_path / f'{file_path.stem}_copy{index + 1}{file_path.suffix}'
                shutil.copy2(str(file_path), str(target_path))

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')