from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader
import shutil



def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        shutil.copytree(str(input_path), str(output_path))
    else:
        shutil.rmtree(str(output_path))
        shutil.copytree(str(input_path), str(output_path))
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    file_list = list(output_path.glob('**/*.pdf'))
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Extracting Text: done'),
        **options
    )

    for file_path in results:
        results.text(f'Extracting Text: {file_path.name}')
        reader = PdfReader(file_path)
        pages = list(reader.pages)
        items = []
        for page in pages:
            item = page.extract_text()
            items.append(item)
        text = '\n'.join(items)
        target_path = file_path.with_suffix('.txt')
        with open(target_path, 'w') as file:
            file.write(text)
        file_path.unlink()

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')