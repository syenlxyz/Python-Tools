from datetime import datetime
from docx import Document
from pathlib import Path

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = list(input_path.glob('*.docx'))
    for file_path in file_list:
        num_page = get_num_page(file_path)
        print(f'{file_path.name}: {num_page}')

def get_num_page(file_path):
    document  = Document(file_path)
    paragraphs = list(document.paragraphs)
    results = []
    for paragraph in paragraphs:
        result = paragraph.contains_page_break
        results.append(result)
    num_page = sum(results) + 1
    return num_page

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')