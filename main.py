import sys
import os
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QFileDialog, QApplication
from PyQt5 import uic
from PyQt5.QtGui import QIcon
from pdfrw import PdfReader, PdfWriter, PdfDict
from shutil import rmtree
from tempfile import TemporaryDirectory, mkdtemp
from configparser import ConfigParser
from datetime import datetime
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import win32com.client as winc
from PIL import Image
from numba import jit


def convert_cover_to_pdf(subdir_Iterator, temp_dir_path) -> None:
    '''개별 표지를 pdf로 변환해 임시 폴더에 저장합니다.'''
    cover_file_path = os.path.join(subdir_Iterator.path, cover_file_name)
    if os.path.isfile(cover_file_path):
        if '.docx' in cover_file_name:
            convert_docx_to_pdf(cover_file_path, os.path.join(temp_dir_path, '0.pdf'))
        elif '.jpg' or '.jpeg' or'.png' or '.tiff' in cover_file_name:
            convert_imgs_to_pdf((cover_file_path), os.path.join(temp_dir_path, '0.pdf'))
        else:
            pass
    else:
        pass
    

def make_filelist_from_subdir_info(subdir_iterator) -> list:
    '''하위폴더 scandir iterator를 받아와 병합에 사용될 해당 하위폴더의 파일들을 추려내고 리스트로 반환합니다.'''
    subdir_filelist = []
    with os.scandir(subdir_iterator.path) as entries:
        subdir_filelist = [entry for entry in entries if entry.is_file() and entry.name != cover_file_name and entry.name not in files_to_except]
        subdir_filelist.sort(key=lambda x: x.name)
    return subdir_filelist

def convert_to_pdf(file_list, temp_dir_path):
    '''변환할 파일들의 scandir iterator들이 있는 리스트를 받아 변환 작업을 수행합니다.'''
    file_list = tuple(file_list)
    if mode == 1:  #docx 병합 모드
        for i, entry in enumerate(file_list):
            if entry.name.endswith('.docx'):
                convert_docx_to_pdf(entry.path, os.path.join(temp_dir_path, f'1{i}.pdf'))
    elif mode == 2:  #이미지 파일 병합 모드
        imgs_for_pdf = tuple([entry.path for entry in file_list if entry.name.endswith('.jpg') or entry.name.endswith('.jpeg') or entry.name.endswith('.png') or entry.name.endswith('.tiff')])
        convert_imgs_to_pdf(imgs_for_pdf, os.path.join(temp_dir_path, '1.pdf'))

def convert_docx_to_pdf(docx_path, pdf_path):
    Application = winc.Dispatch("Word.Application")
    docx = Application.Documents.Open(docx_path)
    docx.ExportAsFixedFormat2(pdf_path, 17, OptimizeFor=0) #17 for PDF and 18 for XPS file forma
    Application.Quit()
@jit(cache=True)
def convert_imgs_to_pdf(img_paths_list: tuple, pdf_path):
    img_list = []
    for i, image_path in enumerate(img_paths_list):
        if len(img_paths_list) == 1:
            image_ = Image.open(image_path)
            image = image_.convert('CMYK')
            image.save(pdf_path)
        else:
            if i == 1:
                image_ = Image.open(image_path)
                image_base = image_.convert('CMYK')
            elif i + 1 < len(img_paths_list):
                image_ = Image.open(image_path)
                image = image_.convert('CMYK')
                img_list.append(image)
            elif i + 1 == len(img_paths_list):
                image_ = Image.open(image_path)
                image = image_.convert('CMYK')
                img_list.append(image)
                image_base.save(pdf_path, save_all=True, append_images=img_list)
            else:
                pass

#아래 함수들에서 사용하기 위해 만든 함수 - 경로 내 하위폴더의 [부문명, 참가자 학번, 참가자명]를 리스트로 반환합니다
def prepare_data(path: str) -> list:
    '''해당 path에 존재하는 모든 폴더와 파일들의 메타데이터들을 가져옵니다'''
    categories = [] #부문명
    student_numbers = [] #학번
    entries = [] #참가자명
    subdirs_infos = []
    with os.scandir(path) as metadatas:
        subdirs_infos = [metadata for metadata in metadatas if metadata.is_dir()]
        subdirs_infos.sort(key=lambda x: main_sep.join([x.name.split(main_sep)[1], x.name.split(main_sep)[2]]))

        categories = [metadata.name.split(main_sep)[0].split(sub_sep)[0].strip() for metadata in subdirs_infos]
        student_numbers = [metadata.name.split(main_sep)[1].split(sub_sep)[0].strip() for metadata in subdirs_infos]
        entries = [metadata.name.split(main_sep)[2].split(sub_sep)[0].strip() for metadata in subdirs_infos]          
        return categories, student_numbers, entries, subdirs_infos
        
def get_subdirs_infos(path: str) -> tuple:
    '''선택된 경로에서 사용 가능한 하위폴더들의 ScandirIterator 목록을 리스트로 반환합니다'''
    return prepare_data(path)[3]

#부문명을 문자열로 반환하는 함수
def get_category(path: str) -> str:
    '''부문명을 뽑아냅니다. 부문명은 한 개여야 하며, 여러 개가 감지되었을 경우 에러를 뱉습니다.'''
    category_name = list(set(prepare_data(path)[0]))

    if len(category_name) == 1:
        category_name = category_name[0]
    elif len(category_name) >= 2:
        category_name = 'Error! 감지된 부문명이 여러 개임'
    else:
        category_name = 'Error! 감지된 부문명이 없음'
    
    return category_name

def get_entries_info(path: str) -> list:
    '''참가자 학번과 참가자명 목록을 리스트로 반환합니다
         ex)   [[2022111111, 김뫄뫄], [2022110000, 김솨솨]]'''
    try:
        infos = list(zip(prepare_data(path)[1], prepare_data(path)[2]))
        for i, _ in enumerate(infos):
            infos[i] = list(_)

        return infos
    except:
        return ["Error! 데이터들의 숫자가 맞지 않음"]

def merge_pdfs_in_dir(dir_path_to_read: str, dir_path_to_write: str, pdf_name: str) -> None:
    '''폴더에 있는 pdf 파일들을 알파벳 순서로 정렬하고 병합해 다른 폴더에 저장합니다.'''
    merger = PdfFileMerger()
    entries = os.scandir(dir_path_to_read)
    pdf_paths = [entry for entry in entries]
    pdf_paths.sort(key=lambda x: x.name)
    
    for pdf in pdf_paths:
        merger.append(pdf.path, import_outline=True)
    
    merger.write(os.path.join(dir_path_to_write, pdf_name))
    merger.close()

def resize_pdf(pdfpath_for_read: str, folderpath_for_write: str) -> str:
    '''pdf각 페이지의 사이즈를 params.py에 적힌 사이즈로 바꿔 임시폴더에 저장하고, 해당 파일의 경로를 반환합니다. 
    단, 각 pdf의 메타데이터 내용과 각 페이지의 모든 북마크가 삭제됩니다. 함수 사용 시 주의할 것!'''
    pdf_file = PdfFileReader(open(pdfpath_for_read, 'rb'))
    pages_num = pdf_file.numPages
    writer = PdfFileWriter()
    
    width = size[0] * 2.83464567
    height = size[1] * 2.83464567

    resized_pdf_path = os.path.join(folderpath_for_write, f'resized_{size[0]}x{size[1]}.pdf')

    for i in range(0, pages_num):
        page_to_resize = pdf_file.getPage(i)
        page_to_resize.scaleTo(width, height)
        writer.addPage(page_to_resize)

    resized_pdf = open(resized_pdf_path, "wb")
    writer.write(resized_pdf)
    resized_pdf.close()
    
    return resized_pdf_path

def add_bookmark(pdf_path: str, bookmarks_makers: list, save_pdf_path: str) -> None:
    '''읽어온 pdf 파일의 첫번째 장에 북마크를 추가합니다.'''
    writer = PdfFileWriter()
    pdf_to_add_bookmark = PdfFileReader(pdf_path)
    page_to_bookmark = 0
    to_bookmark_pagenums = []
    bookmarks = []
    pdf_to_add_bookmark_totalpage = pdf_to_add_bookmark.numPages

    bookmarks_makers = tuple(bookmarks_makers)
    for bookmark, totalpage in bookmarks_makers:
        for i, pagenum in enumerate(range(totalpage)):
            if i == 0:
                to_bookmark_pagenums.append(page_to_bookmark)
                bookmarks.append(bookmark)
                page_to_bookmark += totalpage
            else:
                pass
    for pagenum in range(pdf_to_add_bookmark_totalpage):
        page_to_add = pdf_to_add_bookmark.getPage(pagenum)
        if pagenum in to_bookmark_pagenums:
            writer.addPage(page_to_add)
            writer.addBookmark(bookmarks[to_bookmark_pagenums.index(pagenum)], pagenum, parent=None)
        else:
            writer.addPage(page_to_add)

    bookmarked_pdf_path = open(save_pdf_path, 'wb')
    writer.write(bookmarked_pdf_path)
    bookmarked_pdf_path.close()

def resized_ok(pdf_path: str, bookmark: str, save_pdf_path: str) -> list:
    '''사이즈 조정된 pdf 파일의 경로를 받아와 다른 폴더로 저장하고, 추후 북마크를 추가하기 위해 필요한 자료들을 리스트로 반환합니다.
    [북마크(str), 해당 pdf 파일의 전체 페이지 수(int)]'''
    writer = PdfFileWriter()
    pdf = PdfFileReader(pdf_path)
    pages = pdf.numPages
    for i in range(pages):
        writer.addPage(pdf.getPage(i))
    _pdf_path = open(save_pdf_path, 'wb')
    writer.write(_pdf_path)
    _pdf_path.close()
    return bookmark, pages

version = '0.0.0'
year = datetime.now().year
author = "만화사랑 편집위원 도우미"
title = f"만사 작품집 {year}"

#
def set_path(_path: str) -> None:
    '''폴더경로 파라미터 - 프로그램 이곳저곳에서 사용될 파일경로를 저장합니다.
    params.py의 path를 사용자가 지정한 폴더 경로로 바꿉니다'''
    global path
    path = _path
###

def set_files_to_except(_files_to_except: str) -> None:
    '''해당 파라미터로 입력된 파일명들이 폴더 내에 존재할 경우 해당 파일은 변환 및 병합에 사용하지 않고 넘어갑니다'''
    global files_to_except
    files_to_except = [file_to_except.strip() for file_to_except in _files_to_except.strip().split(',')]
###    

def set_size(width: str, height: str) -> None:
    '''pdf의 사이즈를 mm로 입력합니다'''
    global size
    size = int(width), int(height)
###

def set_seps(_main_sep: str, _sub_sep: str) -> None:
    '''프로그램에서 사용하는 구분자 설정 - 폴더명을 파싱할 때 main_sep을 기준으로 쪼개고, 쪼개진 파츠를 사용할 때 sub_sep 이후에 적힌 문자열은 없는 문자열 취급합니다.'''
    global main_sep, sub_sep
    main_sep = _main_sep 
    sub_sep = _sub_sep
###

def set_cover_file_name(_cover_file_name: str) -> None:
    '''각 작품별 표지로 사용될 개별 표지 파일의 파일명을 설정합니다. 
지원되는 파일 확장자는 다음과 같습니다. => [.docx / .jpg / .png / .tiff]
'''
    global cover_file_name
    cover_file_name = _cover_file_name
###

def set_mode(_mode: int) -> None:
    '''1-2. 모드 설정 파라미터 - pdf 병합 방식을 결정하는 값을 저장합니다. 
    1은 docx 변환 모드, 2는 이미지 파일 변환 모드입니다.'''
    global mode
    mode = _mode
###

######Pdf변환 진행바 구현을 위한 파라미터######
def set_subdir_num(_subdir_num: int) -> None:
    '''진행바 구현을 위해 변환에 사용되는 하위폴더 수를 기록합니다'''
    global subdir_num
    subdir_num = _subdir_num

def set_conversion_step(_step: int) -> None:
    '''진행바 구현을 위해 pdf 변환 과정에서의 스텝을 기록합니다'''
    global step
    step = _step

def set_max_conversion_step(_max_step: int) -> None:
    '''진행바 구현을 위해 최대 스텝 수를 기록합니다'''
    global max_step
    max_step = _max_step

def add_step() -> None:
    '''진행바 구현을 위해 스텝을 가산합니다'''
    global step
    step += 1
######Pdf변환 진행바 구현을 위한 파라미터######

#setting.ini 읽어오기
config = ConfigParser()
config.read('setting.ini', encoding='utf-8')

if config == []:
    with open('setting.ini') as configfile:
        config.add_section("DEFAULT")
        config.set("DEFAULT", 'width(mm)', '182')
        config.set("DEFAULT", 'height(mm)', '254')
        config.set("DEFAULT", 'main separator', '_')
        config.set("DEFAULT", 'sub separator', '+')
        config.set("DEFAULT", 'cover file names', '한 마디.docx')
        config.set("DEFAULT", 'files to except', '왜 3mm 연장해서 원고를 해야 할까.jpg, 작업은 이렇게 해 주세요.txt')
        config.write(configfile)
else:
    if not 'width(mm)' in config:
        config.set("DEFAULT", 'width(mm)', '182')
    if not 'height(mm)' in config:
        config.set("DEFAULT", 'height(mm)', '254')
    if not 'main_separator' in config:
        config.set("DEFAULT", 'main separator', '_')
    if not 'sub_separator' in config:
        config.set("DEFAULT", 'sub separator', '+')
    if not 'cover_file_name' in config:
        config.set("DEFAULT", 'cover file names', '한 마디.docx')
    if not 'files_to_except' in config:
        config.set("DEFAULT", 'files to except', '왜 3mm 연장해서 원고를 해야 할까.jpg, 작업은 이렇게 해 주세요.txt')
    with open('setting.ini', 'w', encoding='utf-8') as configfile:
        config.write(configfile)

form_class = uic.loadUiType("interface.ui")[0]

set_seps(config['DEFAULT']['main separator'], config['DEFAULT']['sub separator'])
set_size(config['DEFAULT']['width(mm)'], config['DEFAULT']['height(mm)'])
set_cover_file_name(config['DEFAULT']['cover file names'])
set_files_to_except(config['DEFAULT']['files to except'])

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon('icon.png'))

        self.select_folder.clicked.connect(self._select_folder)
        self.save_csv.clicked.connect(self._save_csv)
        self.run_convert.clicked.connect(self.make_pdf)
        self.mode_docx.clicked.connect(self.setmode_one)
        self.mode_img.clicked.connect(self.setmode_two)
        self.show_cred.triggered.connect(self.show_info_popup)
        self.initUi()

    def initUi(self):
        self.mode_docx.setCheckable(False)
        self.mode_img.setCheckable(False)
        self.save_csv.setDisabled(True)
        self.run_convert.setDisabled(True)
        self.label_saved_message.hide()
        self.progress.hide()
        self.run_convert.setText('pdf로 변환 시작!')

    def show_info_popup(self) -> None:
        content = f"""< {author} >

버전: {version}
만든 사람: 2022 연세대학교 만화사랑 김나현
아이콘 그린 사람: 2022 연세대학교 만화사랑 신하영"""
        msg = QMessageBox()
        msg.setWindowTitle('프로그램 정보')
        msg.setText(content)
        msg.exec_()

    def _select_folder(self):
        global temp_entries
        global temp_category
        self.initUi()

        dir_path = QFileDialog.getExistingDirectory(self, '불러올 폴더를 선택하세요')
        set_path(dir_path)

        try:
            _temp_entries = [','.join(_) for _ in get_entries_info(path)]

            is_ok_to_save_csv = True
            temp_category = get_category(path)
            temp_entries = '\n'.join(_temp_entries)
            if len(temp_entries) == 0 or len(get_category(path)) == 0:
                temp_entries = 'Error! 존재하지 않음'
                is_ok_to_save_csv = False
            if len(get_category(path)) == 0:
                temp_category = 'Error! 감지된 부문명이 없음'
            
            dir_info = f"부문명: {temp_category}\n\n<참가자 목록>\n{temp_entries}"
            self.show_info.setPlainText(dir_info)
            self.show_path.setText(path)

            if is_ok_to_save_csv:
                self.save_csv.setEnabled(True)
                self.mode_docx.setCheckable(True)
                self.mode_img.setCheckable(True)
        except:
            dir_info = f"부문명: Error! 존재하지 않음\n\n<참가자 목록>\nError! 감지된 부문명이 없음"
            self.show_info.setPlainText(dir_info)
            self.show_path.setText(path)

    def execute_pdf_conversion(self, subdir_infos: list, path_to_write_pdf: str, category_name: str) -> None:
        '''선택한 폴더의 하위 폴더에 대한 scandir iterator들이 있는 리스트를 받아 pdf 변환을 진행합니다.'''
        pdf_name = f'{category_name}_병합본_original.pdf'
        pdf_name_final = f'{category_name}_병합본_resized_{size[0]}x{size[1]}.pdf'
        global to_remove
        to_remove = []
        bookmarks_makers = []

        #QProgressBar 적용을 위해 스텝 수를 기록합니다
        set_conversion_step(0)
        set_subdir_num(len(subdir_infos))
        set_max_conversion_step(subdir_num*2  +  4)

        self.progress.setMaximum(max_step)
        self.progress.setMinimum(0)
        self.progress.setValue(0)
        ###

        with TemporaryDirectory() as temp_dir_path_big:
            for subdir_info in subdir_infos:
                remove_path, bookmarks_maker = self.convert_to_individual_pdf(subdir_info, temp_dir_path_big)
                to_remove.append(remove_path)
                bookmarks_makers.append(bookmarks_maker)
                bookmarks_makers = bookmarks_makers
                
            self.make_progressbar_work(f'{category_name[:4]}..병합 중')
            merge_pdfs_in_dir(temp_dir_path_big, path_to_write_pdf, pdf_name)
            to_remove = tuple(to_remove)
        for temp_path in to_remove:
            self.make_progressbar_work('임시폴더 삭제 중')
            rmtree(temp_path, ignore_errors=True)

        #북마크 추가
        self.make_progressbar_work('책갈피 추가 중')
        add_bookmark(os.path.join(path_to_write_pdf, pdf_name), bookmarks_makers, os.path.join(path_to_write_pdf, pdf_name_final))
        
        #메타데이터 추가
        self.make_progressbar_work('메타데이터 추가 중')
        pdf_reader = PdfReader(os.path.join(path_to_write_pdf, pdf_name_final))
        metadata = PdfDict(Author=author, Title=f'{title} - {category_name}')
        pdf_reader.Info.update(metadata)
        PdfWriter().write(os.path.join(path_to_write_pdf, pdf_name_final), pdf_reader)
        
    def make_progressbar_work(self, message: str) -> None:
        '''진행바 움직이는 함수'''
        add_step()
        self.progress.setValue(step)
        self.run_convert.setText(message)
    
    def convert_to_individual_pdf(self, subdir_Iterator, temp_dir_path_big) -> list:
        '''하위 디렉터리를 하나씩 방문하며 각각의 표지 파일을 pdf로 변환해 임시 폴더에 저장하고, 예외 설정되지 않은 모든 docx 파일 혹은 이미지 파일들을 pdf로 병합합니다. 해당 pdf들을 PdfFileMerger에 넘겨준 후, 임시 폴더를 삭제하는 데에 사용할 주소와 북마크, 페이지 수를 담은 리스트를 반환합니다.[임시폴더경로, 북마크, 해당 개별 pdf의 페이지 수]  개인별로 병합된 pdf파일들이 또 다른 임시폴더에 추가됩니다.''' 
        subdir_name_data = subdir_Iterator.name.split(main_sep)
        _category_name = subdir_name_data[0].split(sub_sep)[0].strip()
        student_number = subdir_name_data[1].split(sub_sep)[0].strip()
        entry_name = subdir_name_data[2].split(sub_sep)[0].strip()

        individual_pdf_name = f'{_category_name}_{student_number}_{entry_name}.pdf'
        bookmark = f'{entry_name}'

        if len(entry_name) > 3:
            message_entry_name = f'{entry_name[:3]}..'
        else:
            message_entry_name = entry_name

        #임시폴더 만들면서 개별 병합된 pdf들이 저장됨
        temp_dir_path = mkdtemp()

        self.make_progressbar_work(f'{message_entry_name} - 개별표지 변환')
        convert_cover_to_pdf(subdir_Iterator, temp_dir_path)

        self.make_progressbar_work(f'{message_entry_name} - 작품파일 병합')
        convert_to_pdf(make_filelist_from_subdir_info(subdir_Iterator), temp_dir_path)
        
        merge_pdfs_in_dir(temp_dir_path, temp_dir_path, individual_pdf_name)
        resized_path = resize_pdf(os.path.join(temp_dir_path, individual_pdf_name), temp_dir_path)
        for_bookmark = resized_ok(resized_path, bookmark, os.path.join(temp_dir_path_big, individual_pdf_name))
        return temp_dir_path, for_bookmark
        
    def _save_csv(self):
        save_csv_path = QFileDialog.getExistingDirectory(self, '파일을 저장할 폴더를 선택하세요', path)
        with open(os.path.join(save_csv_path, f'{temp_category}_참가자_목록.txt'), 'w', encoding='utf-8') as f:
            f.write(f'학번, 참가자명\n{temp_entries}')
        try:
            rmtree(f'{temp_category}_참가자_목록.txt')
        except:
            pass
        self.label_saved_message.show()
    
    def setmode_one(self):
        set_mode(1) #docx 모드
        self.run_convert.setEnabled(True)
        
    def setmode_two(self):
        set_mode(2) #이미지 모드
        self.run_convert.setEnabled(True)

    def make_pdf(self):
        try:
            path_ = path.replace('/', '\\')
            self.run_convert.setDisabled(True)
            self.progress.show()
            self.execute_pdf_conversion(get_subdirs_infos(path_), path_, temp_category)
            self.run_convert.setText('변환 완료^w^')

        except(PermissionError):
            self.show_info.setPlainText(f'Error: 변환 중단됨\n[Permission denied]\n저장하려는 pdf 파일과 동일한 이름인 파일을 os에서 사용하고 있는 것 같습니다.')
            for temp_path in to_remove:
                rmtree(temp_path, ignore_errors=True)
        except:
            for temp_path in to_remove:
                rmtree(temp_path, ignore_errors=True)
def main():
    app = QApplication(sys.argv)
    main_window = WindowClass()
    
    main_window.show()
    sys.exit(app.exec_())

    
if __name__ == "__main__":
    main()
    os.system('pause')
