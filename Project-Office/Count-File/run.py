## Unresolved issue: need better format for displaying result
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
    
    suffix_dict = {}
    for file_path in file_list:
        suffix = file_path.suffix
        suffix_list = list(suffix_dict.keys())
        if suffix not in suffix_list:
            suffix_dict[suffix] = 1
        else:
            suffix_dict[suffix] += 1
    
    print(f'Files: {len(file_list)}')
    print(f'Folders: {len(folder_list)}')
    for folder_path in folder_list:
        file_list = [path for path in folder_path.iterdir() if path.is_file()]
        print(f'  {folder_path.relative_to(input_path)}: {len(file_list)}')
    suffix_list = list(suffix_dict.keys())
    print(f'Extensions: {len(suffix_list)}')
    for suffix in suffix_list:
        print(f'  {suffix}: {suffix_dict[suffix]}')

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')