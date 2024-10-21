from datetime import datetime
from pathlib import Path
from spire.doc import Document

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = []
    path_list = list(input_path.iterdir())
    for path in path_list:
        if path.suffix in ['.doc', '.docx']:
            file_list.append(path)
    
    for file_path in file_list:
        num_page = get_num_page(file_path)
        print(f'{file_path.name}: {num_page}')

def get_num_page(file_path):
    doc = Document(str(file_path))
    properties = doc.BuiltinDocumentProperties
    num_page = properties.PageCount
    return num_page

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')