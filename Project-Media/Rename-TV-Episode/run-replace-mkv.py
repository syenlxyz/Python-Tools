from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from send2trash import send2trash
import re
import shutil
import subprocess

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
        finalize=lambda bar: bar.text('Processing: done'),
        **options
    )
    
    for folder_path in results:
        folder_name = folder_path.name
        results.text(f'Processing: {folder_name}')
        file_list = list(folder_path.glob('*.mkv'))
        for file_path in file_list:
            subprocess.run(f'mkvpropedit "{file_path}" -e info -s title="{file_path.stem}"')
        
        pattern = r'\[(.*?)\]'
        result = re.findall(pattern, folder_name)
        english = result[0]
        chinese = result[1]
        subtitle = result[-3]
        resolution = result[-2]
        
        file_list = list(folder_path.glob('*.mkv'))
        season = get_season(folder_path)
        episodes = get_episodes(file_list)
        
        for index, episode in enumerate(episodes):
            file_path = file_list[index]
            if episode:
                file_name = f'{english}.{chinese}.S{season}E{episode}.{subtitle}.{resolution}'
                target_file = folder_path / f'{file_name}.mkv'
                file_path.rename(target_file)
                
                pattern = file_path.stem.replace('[', '[{}').replace(']', '{}]').replace('{}', '[]')
                caption_list = list(folder_path.glob(f'{pattern}*'))
                for caption_path in caption_list:
                    caption_name = caption_path.name.replace(file_path.stem, target_file.stem)
                    target_caption = folder_path / caption_name
                    caption_path.rename(target_caption)
            else:
                target_file = output_path / file_path.name
                shutil.move(file_path, target_file)

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
    episodes = []
    for file_path in file_list:
        keywords = ['.5', 'OVA']
        for keyword in keywords:
            if keyword in file_path.stem:
                episode = None
                break
        else:
            episode = get_episode(file_path)
        episodes.append(episode)
    
    start = min([int(episode) for episode in episodes if episode])
    if start > 1:
        episodes = [ifelse(episode, str(int(episode) - start + 1), episode) for episode in episodes]
    
    num_digit = max(max([len(episode) for episode in episodes if episode]), 2)
    episodes = [ifelse(episode, episode.zfill(num_digit), episode) for episode in episodes]
    return episodes

def get_episode(file_path):
    patterns = [
        r'\[E\d+\]',
        r'\[\d+(v\d)?\]',
        r'\[\d+\s?END\]',
        r'\[\d+_\d+\]',
        r'-\s?\d+',
        r'\.\d+',
        r'第\d+话',
        r'第\d+話'
    ]
    
    for pattern in patterns:
        result = re.search(pattern, file_path.stem)
        if result:
            text = result.group()
            search = re.search(r'\d+', text)
            episode = search.group()
            return episode

def ifelse(test_expression, x, y):
    if test_expression:
        return x
    else:
        return y

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')