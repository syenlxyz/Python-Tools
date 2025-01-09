from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pypdf import PdfWriter
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
    
    file_list = list(input_path.glob('*.pdf'))
    if file_list:
        num_file = len(file_list)
        num_top = num_file // 2 + 1
        num_bottom = num_file // 2
        top_list = [i * 2 for i in range(num_top)]
        bottom_list = [i * 2 + 1 for i in range(num_bottom)]
        order_list = top_list + bottom_list
        index_list = [order_list.index(value) for value in range(num_file)]
        file_list = [file_list[index] for index in index_list]
        
        results = alive_it(
            file_list, 
            len(file_list), 
            finalize=lambda bar: bar.text(f'Merging PDF in Multiple for {input_path.name}: done'),
            **options
        )
        merger = PdfWriter()
        for file_path in results:
            results.text(f'Merging PDF in Multiple for {input_path.name}: {file_path.name}')
            merger.append(file_path)
        target_path = output_path / input_path.with_suffix('.pdf').name
        merger.write(target_path)
        merger.close()
    
    for folder_path in folder_list:
        file_list = list(folder_path.glob('*.pdf'))
        if file_list:
            num_file = len(file_list)
            num_top = num_file // 2 + 1
            num_bottom = num_file // 2
            top_list = [i * 2 for i in range(num_top)]
            bottom_list = [i * 2 + 1 for i in range(num_bottom)]
            order_list = top_list + bottom_list
            index_list = [order_list.index(value) for value in range(num_file)]
            file_list = [file_list[index] for index in index_list]
            
            results = alive_it(
                file_list, 
                len(file_list), 
                finalize=lambda bar: bar.text(f'Merging PDF in Multiple for {folder_path.name}: done'),
                **options
            )
            merger = PdfWriter()
            for file_path in results:
                results.text(f'Merging PDF in Multiple for {folder_path.name}: {file_path.name}')
                merger.append(file_path)
            target_path = output_path / folder_path.with_suffix('.pdf').name
            merger.write(target_path)
            merger.close()

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')