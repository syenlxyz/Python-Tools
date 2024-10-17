from datetime import datetime
from pathlib import Path

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_count = 0
    folder_count = 0
    file_list = list(input_path.glob('**/*'))
    for file_path in file_list:
        if file_path.is_file():
            file_count += 1
        if file_path.is_dir():
            folder_count += 1
    print(f'File Count: {file_count}')
    print(f'Folder Count: {folder_count}')

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')