from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
import shutil

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        shutil.copytree(str(input_path), str(output_path))
    else:
        shutil.rmtree(str(output_path))
        shutil.copytree(str(input_path), str(output_path))
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = []
    folder_list = []
    path_list = list(output_path.iterdir())
    for path in path_list:
        if path.is_file():
            file_list.append(path)
        if path.is_dir():
            folder_list.append(path)
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text(f'Renaming File for {input_path.name}: done'),
        **options
    )
    for index, file_path in enumerate(results):
        results.text(f'Renaming File {input_path.name}: {file_path.name}')
        num_file = len(file_list)
        num_digit = len(str(num_file))
        prefix = str(index + 1).zfill(num_digit)
        target_path = file_path.parent / f'{prefix}_{file_path.name}'
        file_path.rename(target_path)
    
    for folder_path in folder_list:
        file_list = list(folder_path.iterdir())
        results = alive_it(
            file_list, 
            len(file_list), 
            finalize=lambda bar: bar.text(f'Renaming File for {folder_path.relative_to(output_path)}: done'),
            **options
        )
        for index, file_path in enumerate(results):
            results.text(f'Renaming File for {folder_path.relative_to(output_path)}: {file_path.name}')
            num_file = len(file_list)
            num_digit = len(str(num_file))
            prefix = str(index + 1).zfill(num_digit)
            target_path = file_path.parent / f'{prefix}_{file_path.name}'
            file_path.rename(target_path)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')