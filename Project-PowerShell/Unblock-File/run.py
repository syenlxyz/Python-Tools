from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
import subprocess

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = list(input_path.glob('**/*'))
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
        finalize=lambda bar: bar.text('Processing: done'),
        **options
    )

    for file_path in results:
        results.text(f'Processing: {file_path.name}')
        target_path = str(file_path).replace("'", "''")
        subprocess.run(f"powershell -Command Unblock-File -Path '{target_path}'")

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')