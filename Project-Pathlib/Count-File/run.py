from datetime import datetime
from pathlib import Path

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    folder_list = []
    path_list = list(input_path.glob('**/*'))
    for path in path_list:
        if path.is_dir():
            folder_list.append(path)
    
    data = {}
    for path in path_list:
        parent_path = path.parent
        if parent_path == input_path:
            key = input_path.relative_to(input_path).as_posix()
            if key not in list(data.keys()):
                data[key] = {
                    'file': [],
                    'folder': []
                }
            if path.is_file():
                data[key]['file'].append(path)
            if path.is_dir():
                data[key]['folder'].append(path)
    for folder_path in folder_list:
        for path in path_list:
            parent_path = path.parent
            if parent_path == folder_path:
                key = folder_path.relative_to(input_path).as_posix()
                if key not in list(data.keys()):
                    data[key] = {
                        'file': [],
                        'folder': []
                    }
                if path.is_file():
                    data[key]['file'].append(path)
                if path.is_dir():
                    data[key]['folder'].append(path)
    
    results = sorted(data.items())
    for key, value in results:
        num_file = len(value['file'])
        num_folder = len(value['folder'])
        print(f'{key}: {num_file} Files, {num_folder} Folders')

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')