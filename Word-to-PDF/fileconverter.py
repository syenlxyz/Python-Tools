from pyoffice.filesystem import FileSystem
from pyoffice.fileoptimizer import getExecutionTime
from win32com.client import Dispatch

class FileConverter():
    def __init__(self):
        file_system = FileSystem()
        input_path = file_system.getInputPath()
        output_path = file_system.getOutputPath()
        file_system.getFileCounter(input_path)
        self.convertFiles(input_path, output_path)
    
    def convertFiles(self, input_path, output_path):
        doc_ext = ['.docx',  '.doc']
        xls_ext = ['.xlsx',  '.xls']
        ppt_ext = ['.pptx',  '.ppt']
        other_ext = ['.htm',  '.html',  '.pdf',  '.txt',  '.jpeg',  '.jpg',  '.png']
        print('====================================================================')
        file_list = [path for path in input_path.glob('**/*') if path.is_file()]
        for index, file_path in enumerate(file_list):
            print(f'Converting {file_path.name} ({index + 1}/{len(file_list)})')
            try:
                if file_path.suffix in doc_ext:
                    exec_time = getExecutionTime(self.doc2Pdf, path=file_path, output_path=output_path)
                    print(f'Execution time: {exec_time}')
                elif file_path.suffix in xls_ext:
                    exec_time = getExecutionTime(self.xls2Pdf, path=file_path, output_path=output_path)
                    print(f'Execution time: {exec_time}')
                elif file_path.suffix in ppt_ext:
                    exec_time = getExecutionTime(self.ppt2Pdf, path=file_path, output_path=output_path)
                    print(f'Execution time: {exec_time}')
                elif file_path.suffix in other_ext:
                    exec_time = getExecutionTime(self.other2Pdf, path=file_path, output_path=output_path)
                    print(f'Execution time: {exec_time}')
                else:
                    print(f'{file_path.suffix} is not supported')
            except:
                try:
                    exec_time = getExecutionTime(self.other2Pdf, path=file_path, output_path=output_path)
                    print(f'Execution time: {exec_time}')
                except:
                    print('Conversion failed')
        
    def doc2Pdf(self, path, output_path):
        app = Dispatch('Word.Application')
        app.Visible = False
        doc = app.Documents.Open(str(path), ReadOnly=True)
        output = output_path / path.name.replace(path.suffix, '.pdf')
        doc.SaveAs(str(output), FileFormat=17)
        doc.Close()
        app.Quit()

    def xls2Pdf(self, path, output_path):
        app = Dispatch('Excel.Application')
        app.Visible = False
        xls = app.Workbooks.Open(str(path), ReadOnly=True)
        output = output_path / path.name.replace(path.suffix, '.pdf')
        xls.SaveAs(str(output), FileFormat=17)
        xls.Close()
        app.Quit()

    def ppt2Pdf(self, path, output_path):
        app = Dispatch('Powerpoint.Application')
        ppt = app.Presentations.Open(str(path), ReadOnly=True, WithWindow=False)
        output = output_path / path.name.replace(path.suffix, '.pdf')
        ppt.SaveCopyAs(str(output), FileFormat=32)
        ppt.Close()
        app.Quit()

    def other2Pdf(self, path, output_path):
        app = Dispatch('AcroExch.App')
        app.Hide()
        avDoc = Dispatch('AcroExch.AVDoc')
        avDoc.Open(str(path), '')
        pdDoc = avDoc.GetPDDoc()
        output = output_path / path.name.replace(path.suffix, '.pdf')
        pdDoc.Save(1, output)
        avDoc.Close(True)
        app.MenuItemExecute('Quit')