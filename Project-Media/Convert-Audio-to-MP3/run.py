from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
import shutil
import subprocess

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
    
    file_list = []
    path_list = list(output_path.glob('**/*'))
    suffix_list = ['.aac', '.m4a', '.wav', '.wma', '.webm']
    for path in path_list:
        if path.suffix in suffix_list:
            file_list.append(path)
    
    results = alive_it(
        file_list,
        len(file_list),
        finalize=lambda bar: bar.text('Converting Audio to MP3: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Converting Audio to MP3: {file_path.name}')
        target_path = file_path.with_suffix('mp3')
        subprocess.run(f'ffmpeg -hide_banner -loglevel quiet -i "{file_path} -c copy {target_path}')
        file_path.unlink()

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')