from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = []
    path_list = list(input_path.glob('**/*'))
    for path in path_list:
        if path.suffix in ['.doc', '.docx']:
            file_list.append(path)
    
    wrd = Dispatch('Word.Application')
    wrd.Visible = False
    for file_path in file_list:
        wrd.Documents.Open(str(file_path))
        num_page = wrd.ActiveDocument.ComputeStatistics(Statistic=2)
        wrd.ActiveDocument.Close(SaveChanges=False)
        print(f'{file_path.name}: {num_page}')
    wrd.Quit()

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')