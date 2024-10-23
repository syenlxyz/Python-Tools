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
    
    file_list = []
    path_list = list(input_path.glob('**/*'))
    for path in path_list:
        if path.is_file():
            file_list.append(path)
    
    suffix_list = []
    for file_path in file_list:
        suffix = file_path.suffix
        if suffix not in suffix_list:
            # mkdir
            # copy file
            # append list
            pass
        else:
            # copy file
            pass

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')