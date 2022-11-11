from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from os import scandir, path
import get_datas.params as p
from numba import jit


def merge_pdfs_in_dir(dir_path_to_read: str, dir_path_to_write: str, pdf_name: str) -> None:
    '''폴더에 있는 pdf 파일들을 알파벳 순서로 정렬하고 병합해 다른 폴더에 저장합니다.'''
    merger = PdfFileMerger()
    entries = scandir(dir_path_to_read)
    pdf_paths = [entry for entry in entries]
    pdf_paths.sort(key=lambda x: x.name)
    
    for pdf in pdf_paths:
        merger.append(pdf.path, import_outline=True)
    
    merger.write(path.join(dir_path_to_write, pdf_name))
    merger.close()

@jit(cache=True)
def resize_pdf(pdfpath_for_read: str, folderpath_for_write: str) -> str:
    '''pdf각 페이지의 사이즈를 params.py에 적힌 사이즈로 바꿔 임시폴더에 저장하고, 해당 파일의 경로를 반환합니다. 
    단, 각 pdf의 메타데이터 내용과 각 페이지의 모든 북마크가 삭제됩니다. 함수 사용 시 주의할 것!'''
    pdf_file = PdfFileReader(open(pdfpath_for_read, 'rb'))
    pages_num = pdf_file.numPages
    writer = PdfFileWriter()
    
    width = p.size[0] * 2.83464567
    height = p.size[1] * 2.83464567

    resized_pdf_path = path.join(folderpath_for_write, f'resized_{p.size[0]}x{p.size[1]}.pdf')

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
    print(bookmark, pages)
    return bookmark, pages


