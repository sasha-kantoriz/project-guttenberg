import requests
from datetime import datetime
import pathlib
import fpdf
from bs4 import BeautifulSoup
import re
import openpyxl
from time import sleep
from random import randint


class PDF(fpdf.FPDF):
    def footer(self):
        if self.page_no() != 1:
            # Go to 1.5 cm from bottom
            self.set_y(-15)
            self.add_font("dejavu-sans", style="", fname="assets/DejaVuSans.ttf")
            self.set_font("dejavu-sans", size=8)
            # Print centered page number
            self.cell(0, 10, f"{self.page_no()}", 0, 0, 'C')


def get_latest_published_book_index():
    url = 'https://www.gutenberg.org/ebooks/search/?sort_order=release_date'
    response = requests.get(url, timeout=60)
    html = BeautifulSoup(response.content, features="html.parser")
    latest_book = html.body.find('li', attrs={'class': 'booklink'})
    index = latest_book.a.attrs['href'].split('/')[-1]
    return int(index)

def update_last_index(index):
    with open('index', 'w') as f:
        f.write(str(index))

def get_previous_last_index():
    try:
        return int(open('index').read()) + 1
    except:
        return 1

def generate_book_pdf(folder, _id, title, author, content):
    currentYear, currentMonth = datetime.now().year, datetime.now().month
    pdf_fname = f"{currentYear}_{currentMonth}_{_id}_paperback_interior.pdf"
    pdf = PDF(format=(152.4, 228.6))
    pdf.add_font("dejavu-sans", style="", fname="assets/DejaVuSans.ttf")
    pdf.add_page()
    pdf.set_font("dejavu-sans", size=24)
    lines_num = len(pdf.multi_cell(w=0, align='C', padding=(0, 8), text=f"{title}\n{author}", dry_run=True, output="LINES"))
    if lines_num >= 3:
        padding_top = (228.6 - 24 * (lines_num - 1)) / 2
    else:
        padding_top = (228.6 - 24 * (lines_num)) / 2
    pdf.multi_cell(w=0, align='C', padding=(padding_top, 8, 0), text=f"{title}\n{author}")
    pdf.add_page()
    pdf.set_font("dejavu-sans", size=12)
    pdf.multi_cell(w=0, align='J', padding=8, text=content)
    pages = pdf.page_no()
    pdf.output(f"{folder}/{pdf_fname}")
    return pdf_fname, pages

def get_books(datestamp, start, end):
    update_index_flag = True
    try:
        wb = openpyxl.load_workbook('Project Guttenberg.xlsx')
    except:
        wb = openpyxl.Workbook()
    try:
        del wb['Sheet']
    except:
        pass
    ws = wb.create_sheet(datestamp)
    ws.append(["Book ID", "Plain text URL", "Title", "Language", "Author", "Translator", "Illustrator", "Pages num", "PDF file name"])
    try:
        for i in range(start, end + 1):
            sleep(randint(1, 3))
            book_url = f'https://www.gutenberg.org/ebooks/{i}'
            book_txt_url = f'{book_url}.txt.utf-8'
            book_txt = requests.get(book_txt_url, timeout=60, headers={'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 6.1; zh-CN) AppleWebKit/533+ (KHTML, like Gecko)'}).content.decode('utf-8')
            #
            book_author = re.search(r"Author: (.*)\n", book_txt)
            book_author = book_author.groups()[0] if book_author else ""
            book_language = re.search(r"Language: (.*)\n", book_txt)
            book_language = book_language.groups()[0] if book_language else ""
            book_translator = re.search(r"Translator: (.*)\n", book_txt)
            book_translator = book_translator.groups()[0] if book_translator else ""
            book_illustrator = re.search(r"Illustrator: (.*)\n", book_txt)
            book_illustrator = book_illustrator.groups()[0] if book_illustrator else ""
            book_title = re.search(r"Title: (.*)\n", book_txt)
            book_title = book_title.groups()[0] if book_title else ""
            book_content_start_index = re.search(r"\*\*\* START OF THE PROJECT GUTENBERG .* \*\*\*", book_txt)
            book_content_start_index = book_content_start_index.end() if book_content_start_index else 0
            book_content_end_index = re.search(r"\*\*\* END OF THE PROJECT GUTENBERG .* \*\*\*", book_txt)
            book_content_end_index = book_content_end_index.start() if book_content_end_index else -1
            book_txt = book_txt[book_content_start_index:book_content_end_index].replace('\r\n\r\n', '_____').replace('\r\n', '').replace('____', '\r\n\r\n').replace('\n\n', '_____').replace('\n', '').replace('____', '\n\n').replace('_', '')
            book_fname, pages_num = generate_book_pdf(datestamp, i, book_title, book_author, book_txt)
            #
            ws.append([i, book_txt_url, book_title, book_language, book_author, book_translator, book_illustrator, pages_num, book_fname])
    except KeyboardInterrupt:
        update_index_flag = False
    except Exception as e:
        print(e)
        update_index_flag = False
        update_last_index(i)
    finally:
        wb.save('Project Guttenberg.xlsx')
        # update last published book index
        if update_index_flag:
            update_last_index(end_index)

if __name__ == '__main__':
    # create PDFs output folder
    datestamp = datetime.now().strftime('%Y-%B-%d')
    pathlib.Path(datestamp).mkdir(parents=True, exist_ok=True)
    # get last published book index
    start_index, end_index = get_previous_last_index(), get_latest_published_book_index()
    #
    get_books(datestamp, start_index, end_index)
