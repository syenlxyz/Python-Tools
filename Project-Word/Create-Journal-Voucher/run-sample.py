from datetime import datetime
from docx import Document
from pathlib import Path
import pandas as pd
import re

def run():
    template_path = Path.cwd() / 'template.docx'
    columns = []
    doc = Document(template_path)
    table = doc.tables[0]
    for row in table.rows:
        for item in row.cells:
            pattern = r'\$\{(.*?)\}'
            result = re.findall(pattern, item.text)
            if result:
                text = result[0]
                if text not in columns:
                    columns.append(text)
    df = pd.DataFrame(columns=columns)
    df.to_excel('sample.xlsx', index=False)

if __name__ == '__main__':
    print(f'Running {Path(__file__).parent.name}')
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')