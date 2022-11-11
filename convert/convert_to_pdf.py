
import get_datas.params as p
from os import path, scandir
import win32com.client
from PIL import Image
from numba import jit

def convert_cover_to_pdf(subdir_Iterator, temp_dir_path) -> None:
    '''개별 표지를 pdf로 변환해 임시 폴더에 저장합니다.'''
    cover_file_path = path.join(subdir_Iterator.path, p.cover_file_name)
    if '.docx' in p.cover_file_name:
        convert_docx_to_pdf(cover_file_path, path.join(temp_dir_path, '0.pdf'))
    elif '.jpg' or '.jpeg' or'.png' or '.tiff' in p.cover_file_name:
        convert_imgs_to_pdf((cover_file_path), path.join(temp_dir_path, '0.pdf'))
    else:
        pass

def make_filelist_from_subdir_info(subdir_iterator) -> list:
    '''하위폴더 scandir iterator를 받아와 병합에 사용될 해당 하위폴더의 파일들을 추려내고 리스트로 반환합니다.'''
    subdir_filelist = []
    with scandir(subdir_iterator.path) as entries:
        subdir_filelist = [entry for entry in entries if entry.is_file() and entry.name != p.cover_file_name and entry.name not in p.files_to_except]
        subdir_filelist.sort(key=lambda x: x.name)
    return subdir_filelist

def convert_to_pdf(file_list, temp_dir_path):
    '''변환할 파일들의 scandir iterator들이 있는 리스트를 받아 변환 작업을 수행합니다.'''
    file_list = tuple(file_list)
    if p.mode == 1:  #docx 병합 모드
        for i, entry in enumerate(file_list):
            if entry.name.endswith('.docx'):
                convert_docx_to_pdf(entry.path, path.join(temp_dir_path, f'1{i}.pdf'))
    elif p.mode == 2:  #이미지 파일 병합 모드
        imgs_for_pdf = tuple([entry.path for entry in file_list if entry.name.endswith('.jpg') or entry.name.endswith('.jpeg') or entry.name.endswith('.png') or entry.name.endswith('.tiff')])
        convert_imgs_to_pdf(imgs_for_pdf, path.join(temp_dir_path, '1.pdf'))

def convert_docx_to_pdf(docx_path, pdf_path):
    Application = win32com.client.Dispatch("Word.Application")
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


