from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from send2trash import send2trash

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
        return None
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        output_path.mkdir()
    else:
        send2trash(output_path)
        output_path.mkdir()
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    folder_list = []
    path_list = list(input_path.iterdir())
    for path in path_list:
        if path.is_dir():
            folder_list.append(path)
    
    for folder_path in folder_list:
        file_list = list(folder_path.glob('*.pdf'))
        results = alive_it(
            file_list,
            len(file_list),
            finalize=lambda bar: bar.text(f'Processing {folder_path.name}: done'),
            **options
        )
        
        for file_path in results:
            results.text(f'Processing {folder_path.name}: {file_path.name}')
            password = folder_path.name
            reader = PdfReader(file_path, password=password)
            writer = PdfWriter()
            pages = list(reader.pages)
            for page in pages:
                writer.add_page(page)
            target_path = output_path / folder_path.name / file_path.name
            with open(target_path, 'wb') as file:
                writer.write(file)

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')