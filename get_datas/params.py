from datetime import datetime


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