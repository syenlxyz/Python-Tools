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
    
    folder_list = list(input_path.iterdir())
    results = alive_it(
        folder_list,
        len(folder_list),
        finalize=lambda bar: bar.text('Renaming TV Subtitle: done'),
        **options
    )
    
    for folder_path in results:
        folder_name = folder_path.name
        results.text(f'Renaming TV Subtitle: {folder_name}')
        file_list = list(folder_path.iterdir())
        for file_path in file_list:
            file_name = file_path.name.replace('_', '-')
            new_path = output_path / f'{folder_name}.{file_name}'
            shutil.copy2(file_path, new_path)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')