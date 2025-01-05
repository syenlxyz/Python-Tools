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
        finalize=lambda bar: bar.text('Compressing Files: done'),
        **options
    )
    
    for folder_path in results:
        results.text(f'Compressing Files: {folder_path.name}')
        target_path = output_path / folder_path.name
        shutil.make_archive(str(target_path), 'zip', str(folder_path))

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')