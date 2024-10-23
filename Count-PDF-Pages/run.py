from datetime import datetime
from pathlib import Path
from pypdf import PdfReader

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = list(input_path.glob('**/*.pdf'))
    for file_path in file_list:
        reader = PdfReader(file_path)
        num_page = reader.get_num_pages()
        print(f'{file_path.relative_to(input_path)}: {num_page}')

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')