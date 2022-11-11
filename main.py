import sys
from os import path
import os
from get_datas.get_info import get_category, get_entries_info, get_subdirs_infos
from merge.merger import merge_pdfs_in_dir, add_bookmark, resize_pdf, resized_ok
from convert.convert_to_pdf import convert_cover_to_pdf, convert_to_pdf, make_filelist_from_subdir_info
import get_datas.params as p
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QIcon
from pdfrw import PdfReader, PdfWriter, PdfDict
from shutil import rmtree
from tempfile import TemporaryDirectory, mkdtemp
from configparser import ConfigParser

#setting.ini 읽어오기
config = ConfigParser()
config.read('setting.ini', encoding='utf-8')

if config == []:
    with open('setting.ini') as configfile:
        config.add_section("DEFAULT")
        config.set("DEFAULT", 'width(mm)', '182')
        config.set("DEFAULT", 'height(mm)', '254')
        config.set("DEFAULT", 'Main seperator', '_')
        config.set("DEFAULT", 'Sub seperator', '+')
        config.set("DEFAULT", 'Cover file name', '한 마디.docx')
        config.set("DEFAULT", 'Files to except', '왜 3mm 연장해서 원고를 해야 할까.jpg, 작업은 이렇게 해 주세요.txt')
        config.write(configfile)
else:
    if not 'width(mm)' in config:
        config.set("DEFAULT", 'width(mm)', '182')
    if not 'height(mm)' in config:
        config.set("DEFAULT", 'height(mm)', '254')
    if not 'main_seperator' in config:
        config.set("DEFAULT", 'main seperator', '_')
    if not 'sub_seperator' in config:
        config.set("DEFAULT", 'sub seperator', '+')
    if not 'cover_file_name' in config:
        config.set("DEFAULT", 'cover file name', '한 마디.docx')
    if not 'files_to_except' in config:
        config.set("DEFAULT", 'files to except', '왜 3mm 연장해서 원고를 해야 할까.jpg, 작업은 이렇게 해 주세요.txt')
    with open('setting.ini', 'w', encoding='utf-8') as configfile:
        config.write(configfile)



form_class = uic.loadUiType("interface.ui")[0]

p.set_seps(config['DEFAULT']['main seperator'], config['DEFAULT']['sub seperator'])
p.set_size(config['DEFAULT']['width(mm)'], config['DEFAULT']['height(mm)'])
p.set_cover_file_name(config['DEFAULT']['cover file name'])
p.set_files_to_except(config['DEFAULT']['files to except'])
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
        content = f"""< {p.author} >

버전: {p.version}
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
        p.set_path(dir_path)

        try:
            _temp_entries = [','.join(_) for _ in get_entries_info(p.path)]

            is_ok_to_save_csv = True
            temp_category = get_category(p.path)
            temp_entries = '\n'.join(_temp_entries)
            if len(temp_entries) == 0 or len(get_category(p.path)) == 0:
                temp_entries = 'Error! 존재하지 않음'
                is_ok_to_save_csv = False
            if len(get_category(p.path)) == 0:
                temp_category = 'Error! 감지된 부문명이 없음'
            
            dir_info = f"부문명: {temp_category}\n\n<참가자 목록>\n{temp_entries}"
            self.show_info.setPlainText(dir_info)
            self.show_path.setText(p.path)

            if is_ok_to_save_csv:
                self.save_csv.setEnabled(True)
                self.mode_docx.setCheckable(True)
                self.mode_img.setCheckable(True)
        except:
            dir_info = f"부문명: Error! 존재하지 않음\n\n<참가자 목록>\nError! 감지된 부문명이 없음"
            self.show_info.setPlainText(dir_info)
            self.show_path.setText(p.path)

    
    def execute_pdf_conversion(self, subdir_infos: list, path_to_write_pdf: str, category_name: str) -> None:
        '''선택한 폴더의 하위 폴더에 대한 scandir iterator들이 있는 리스트를 받아 pdf 변환을 진행합니다.'''
        pdf_name = f'{category_name}_병합본_original.pdf'
        pdf_name_final = f'{category_name}_병합본_resized_{p.size[0]}x{p.size[1]}.pdf'
        global to_remove
        to_remove = []
        bookmarks_makers = []

        #QProgressBar 적용을 위해 스텝 수를 기록합니다
        p.set_conversion_step(0)
        p.set_subdir_num(len(subdir_infos))
        p.set_max_conversion_step(p.subdir_num*2  +  4)

        self.progress.setMaximum(p.max_step)
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
        add_bookmark(path.join(path_to_write_pdf, pdf_name), bookmarks_makers, path.join(path_to_write_pdf, pdf_name_final))
        
        #메타데이터 추가
        self.make_progressbar_work('메타데이터 추가 중')
        pdf_reader = PdfReader(path.join(path_to_write_pdf, pdf_name_final))
        metadata = PdfDict(Author=p.author, Title=f'{p.title} - {category_name}')
        pdf_reader.Info.update(metadata)
        PdfWriter().write(path.join(path_to_write_pdf, pdf_name_final), pdf_reader)
        

    def make_progressbar_work(self, message: str) -> None:
        '''진행바 움직이는 함수'''
        p.add_step()
        self.progress.setValue(p.step)
        self.run_convert.setText(message)
    
    def convert_to_individual_pdf(self, subdir_Iterator, temp_dir_path_big) -> list:
        '''하위 디렉터리를 하나씩 방문하며 각각의 표지 파일을 pdf로 변환해 임시 폴더에 저장하고, 예외 설정되지 않은 모든 docx 파일 혹은 이미지 파일들을 pdf로 병합합니다. 해당 pdf들을 PdfFileMerger에 넘겨준 후, 임시 폴더를 삭제하는 데에 사용할 주소와 북마크, 페이지 수를 담은 리스트를 반환합니다.[임시폴더경로, 북마크, 해당 개별 pdf의 페이지 수]  개인별로 병합된 pdf파일들이 또 다른 임시폴더에 추가됩니다.''' 
        subdir_name_data = subdir_Iterator.name.split(p.main_sep)
        _category_name = subdir_name_data[0].split(p.sub_sep)[0].strip()
        student_number = subdir_name_data[1].split(p.sub_sep)[0].strip()
        entry_name = subdir_name_data[2].split(p.sub_sep)[0].strip()

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
        resized_path = resize_pdf(path.join(temp_dir_path, individual_pdf_name), temp_dir_path)
        for_bookmark = resized_ok(resized_path, bookmark, path.join(temp_dir_path_big, individual_pdf_name))
        return temp_dir_path, for_bookmark
        
    def _save_csv(self):
        save_csv_path = QFileDialog.getExistingDirectory(self, '파일을 저장할 폴더를 선택하세요', p.path)
        with open(path.join(save_csv_path, f'{temp_category}_참가자_목록.txt'), 'w', encoding='utf-8') as f:
            f.write(f'학번, 참가자명\n{temp_entries}')
        self.label_saved_message.show()
    
    def setmode_one(self):
        p.set_mode(1) #docx 모드
        self.run_convert.setEnabled(True)
        

    def setmode_two(self):
        p.set_mode(2) #이미지 모드
        self.run_convert.setEnabled(True)

    def make_pdf(self):
        try:
            path_ = p.path.replace('/', '\\')
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
