from datetime import datetime
from pathlib import Path
import subprocess

def run():
    url_list = []
    while True:
        url = input('Paste link here (or press ENTER to continue): ')
        if url:
            url_list.append(url)
        else:
            break
    
    for url in url_list:
        print(f'Processing URL: {url}')
        subprocess.run(f'anyflip-downloader {url}')

if __name__ == '__main__':
    package = Path(__file__).parent
    module = Path(__file__)
    print(f'Running {package.name}/{module.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')