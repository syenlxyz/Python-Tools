from alive_progress import alive_it
from datetime import datetime
from mutagen import File
from pathlib import Path
from send2trash import send2trash
import re
import shutil

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
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
    
    folder_list = [path for path in input_path.iterdir() if path.is_dir()]
    results = alive_it(
        folder_list,
        len(folder_list),
        finalize=lambda bar: bar.text('Renaming TV Episode: done'),
        **options
    )
    
    for folder_path in results:
        folder_name = folder_path.name
        results.text(f'Renaming TV Episode: {folder_name}')
        file_list = list(folder_path.glob('*.mp4'))
        for file_path in file_list:
            file_name = file_path.name
            if len(file_name.split('.')) != 6:
                file = File(file_path, easy=True)
                file['title'] = file_path.stem
                file.save()
            else:
                file = File(file_path, easy=True)
                title = file['title']
                target_path = file_path.parent / f'{title}.mp4'
                file_path.rename(target_path)
        
        pattern = r'\[(.*?)\]'
        result = re.findall(pattern, folder_name)
        english = result[0]
        chinese = result[1]
        subtitle = result[-3]
        resolution = result[-2]
        
        file_list = list(folder_path.glob('*.mp4'))
        season = get_season(folder_path)
        episodes = get_episodes(file_list)
        
        for index, episode in enumerate(episodes):
            file_path = file_list[index]
            if episode:
                file_name = f'{english}.{chinese}.S{season}E{episode}.{subtitle}.{resolution}.mp4'
                target = folder_path / file_name
                file_path.rename(target)
            else:
                target = output_path / file_path.name
                shutil.move(file_path, target)

def get_season(folder_path):
    num_digit = 2
    pattern = r'\[S\d+\]'
    result = re.search(pattern, folder_path.name)
    if result:
        text = result.group()
        search = re.search(r'\d+', text)
        season = search.group().zfill(num_digit)
    else:
        season = '1'.zfill(num_digit)
    return season

def get_episodes(file_list):
    num_file = len(file_list)
    num_digit = max(len(str(num_file + 1)), 2)
    episodes = [str(index + 1).zfill(num_digit) for index in range(num_file)]
    return episodes

def ifelse(test_expression, x, y):
    if test_expression:
        return x
    else:
        return y

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')