import os
import get_datas.params as p
from numba import jit

#아래 함수들에서 사용하기 위해 만든 함수 - 경로 내 하위폴더의 [부문명, 참가자 학번, 참가자명]를 리스트로 반환합니다
def prepare_data(path: str) -> list:
    '''해당 path에 존재하는 모든 폴더와 파일들의 메타데이터들을 가져옵니다'''
    categories = [] #부문명
    student_numbers = [] #학번
    entries = [] #참가자명
    subdirs_infos = []
    with os.scandir(path) as metadatas:
        subdirs_infos = [metadata for metadata in metadatas if metadata.is_dir()]
        subdirs_infos.sort(key=lambda x: p.main_sep.join([x.name.split(p.main_sep)[1], x.name.split(p.main_sep)[2]]))

        categories = [metadata.name.split(p.main_sep)[0].split(p.sub_sep)[0].strip() for metadata in subdirs_infos]
        student_numbers = [metadata.name.split(p.main_sep)[1].split(p.sub_sep)[0].strip() for metadata in subdirs_infos]
        entries = [metadata.name.split(p.main_sep)[2].split(p.sub_sep)[0].strip() for metadata in subdirs_infos]          
        return categories, student_numbers, entries, subdirs_infos
        
        
def get_subdirs_infos(path: str) -> tuple:
    '''선택된 경로에서 사용 가능한 하위폴더들의 ScandirIterator 목록을 리스트로 반환합니다'''
    return prepare_data(path)[3]


#부문명을 문자열로 반환하는 함수

@jit(cache=True)
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



