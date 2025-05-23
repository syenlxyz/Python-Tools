from alive_progress import alive_it
from datetime import datetime
from pathlib import Path

def run():
    input_path = Path.cwd()
    output_path = Path.cwd() / 'README.md'
    folder_list = list(input_path.glob('Project-*'))
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    lines = []
    lines.append('# Python-Tools')
    for index, folder_path in enumerate(folder_list):
        lines.append(f'{index + 1}. {folder_path.name}')
        
        file_list = list(folder_path.iterdir())
        results = alive_it(
            file_list, 
            len(file_list), 
            finalize=lambda bar: bar.text(f'Creating URL for {folder_path.name}: done'),
            **options
        )
        for file_path in results:
            results.text(f'Creating URL for {folder_path.name}: {file_path.name}')
            base_url = 'https://github.com/syenlxyz/Python-Tools/tree/main'
            target_url = '/'.join([base_url, folder_path.name, file_path.name])
            lines.append(f'    * [{file_path.name}]({target_url})')
    
    text = '\n'.join(lines)
    with open(output_path, 'w') as file:
        file.write(text)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')