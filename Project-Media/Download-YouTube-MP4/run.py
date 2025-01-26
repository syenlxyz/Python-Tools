from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from pytubefix import YouTube, Playlist
from send2trash import send2trash
from urllib.parse import urlparse, parse_qs
import json
import subprocess

def run():
    output_path = Path.cwd() / 'output'
    if not output_path.is_dir():
        output_path.mkdir()
    
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    playlist = get_playlist()
    if not playlist:
        return None
    
    results = alive_it(
        playlist,
        len(playlist),
        finalize=lambda bar: bar.text('Downloading YouTube MP4: done'),
        **options
    )
    
    for url in results:
        results.text(f'Downloading YouTube MP4: {url}')
        yt = YouTube(url, use_po_token=True, po_token_verifier=po_token_verifier)
        
        video_file = Path(
            yt.streams
            .filter(only_video=True)
            .order_by('resolution')
            .desc()
            .first()
            .download(output_path)
        )
        video_file = video_file.replace(output_path / f'{video_file.stem} - Video Only{video_file.suffix}')
        
        audio_file = Path(
            yt.streams
            .filter(only_audio=True)
            .order_by('bitrate')
            .desc()
            .first()
            .download(output_path)
        )
        audio_file = audio_file.replace(output_path / f'{audio_file.stem} - Audio Only{audio_file.suffix}')
        
        file_name = video_file.stem.replace(' - Video Only', '')
        output_file = output_path / f'{file_name}.mp4'
        subprocess.run(f'ffmpeg -hide_banner -loglevel quiet -i "{video_file}" -i "{audio_file}" "{output_file}"')
        
        send2trash(video_file)
        send2trash(audio_file)

def get_playlist():
    playlist = []
    while True:
        url = input('Paste link here (or press ENTER to continue): ')
        if not url:
            break
        result = urlparse(url)
        netloc = result.netloc
        if 'youtu.be' == netloc:
            playlist.append(url)
        elif 'youtube.com' in netloc:
            query = result.query
            params = parse_qs(query)
            keys = list(params.keys())
            if 'list' in keys:
                p = Playlist(url)
                playlist.extend(p.video_urls)
            elif 'v' in keys:
                playlist.append(url)
            else:
                print('Invalid URL. Please try again.')
        else:
            print('Invalid URL. Please try again.')
    return playlist

def po_token_verifier():
    result = subprocess.run('node script.js', capture_output=True)
    data = json.loads(result.stdout)
    po_token = tuple(data.values())
    return po_token

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')