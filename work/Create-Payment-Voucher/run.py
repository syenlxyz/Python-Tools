from alive_progress import alive_it
from datetime import datetime
from docx import Document
from pathlib import Path
import pandas as pd
import re
import shutil

num_under_20 = {
    0: '',
    1: 'One',
    2: 'Two',
    3: 'Three',
    4: 'Four',
    5: 'Five',
    6: 'Six',
    7: 'Seven',
    8: 'Eight',
    9: 'Nine',
    10: 'Ten',
    11: 'Eleven',
    12: 'Twelve',
    13: 'Thirteen',
    14: 'Fourteen',
    15: 'Fifteen',
    16: 'Sixteen',
    17: 'Seventeen',
    18: 'Eighteen',
    19: 'Nineteen'
}

num_under_100 = {
    2: 'Twenty',
    3: 'Thirty',
    4: 'Forty',
    5: 'Fifty',
    6: 'Sixty',
    7: 'Seventy',
    8: 'Eighty',
    9: 'Ninety'
}

num_above_1000 = {
    1: 'Thousand',
    2: 'Million',
    3: 'Biliion',
    4: 'Trillion'
}

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

def num_to_word(num):
    decimal = round(num - int(num), 2)
    if decimal:
        return ' '.join([num_to_word(int(num)), 'and Cents', num_to_word(int(decimal * 100))])
    num_digit = len(str(num))
    power = (num_digit - 1) // 3
    if num < 20:
        return num_under_20[num]
    elif num < 100:
        return ' '.join([num_under_100[num // 10], num_to_word(num % 10)])
    elif num < 1000:
        return ' '.join([num_to_word(num // 100), 'Hundred', num_to_word(num % 100)])
    elif power < 5:
        return ' '.join([num_to_word(num // 1000 ** power), num_above_1000[power], num_to_word(num % 1000 ** power)])

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')