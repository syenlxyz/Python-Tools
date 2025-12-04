from datetime import datetime
from pathlib import Path

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = []
    folder_list = []
    path_list = list(input_path.glob('**/*'))
    for path in path_list:
        if path.is_file():
            file_list.append(path)
        if path.is_dir():
            folder_list.append(path)
    
    results = sorted(path_list)
    for index, result in enumerate(results):
        num_path = len(results)
        num_digit = len(str(num_path))
        index = str(index+1).zfill(num_digit)
        target_path = result.relative_to(input_path).as_posix()
        print(f'{index}: {target_path}')

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')