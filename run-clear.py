from datetime import datetime
from pathlib import Path

def run():
    path_list = list(Path.cwd().glob('**/*'))
    for path in path_list:
        name_list = ['input', 'output']
        if path.is_dir() and path.name in name_list:
            print(path)
            #shutil.rmtree(str(path))

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')