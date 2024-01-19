import requests
from datetime import datetime
import pathlib
import fpdf
from bs4 import BeautifulSoup
import re
import openpyxl
from time import sleep
import sys
from random import randint
import argparse
from openai import OpenAI


client = OpenAI()

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

def generate_book_pdfs(folder, _id, title, author, description, content):
    currentYear, currentMonth = datetime.now().year, datetime.now().month
    interior_pdf_fname, cover_pdf_fname = f"{currentYear}_{currentMonth}_{_id}_paperback_interior.pdf", f"{currentYear}_{currentMonth}_{_id}_paperback_cover.pdf"
    pdf = PDF(format=(152.4, 228.6))
    pdf.add_font("dejavu-sans", style="", fname="assets/DejaVuSans.ttf")
    pdf.add_page()
    pdf.set_font("dejavu-sans", size=24)
    lines_num = len(pdf.multi_cell(w=0, align='C', padding=(0, 8), text=f"{title}\n\n{author}", dry_run=True, output="LINES"))
    if lines_num >= 3:
        padding_top = (228.6 - 24 * (lines_num - 1)) / 2
    else:
        padding_top = (228.6 - 24 * (lines_num)) / 2
    pdf.multi_cell(w=0, align='C', padding=(padding_top, 8, 0), text=f"{title}\n\n{author}")
    pdf.add_page()
    pdf.set_font("dejavu-sans", size=12)
    pdf.multi_cell(w=0, align='J', padding=8, text=content)
    pages = pdf.page_no()
    if pages >= 24 and pages <= 828:
        pdf.output(f"{folder}/{interior_pdf_fname}")
        #
        cover_width, cover_height = 152.4 * 2 + pages * 0.05720 + 3.175 * 2, 234.95
        pdf = fpdf.FPDF(format=(cover_width, cover_height))
        pdf.add_font('dejavu-sans', style="", fname="assets/DejaVuSans.ttf")
        pdf.add_page()
        pdf.set_fill_color(r=250,g=249,b=222)
        pdf.rect(h=pdf.h, w=pdf.w, x=0, y=0, style="DF")
        cols = pdf.text_columns(ncols=2, gutter=pages*0.05720 + 1.588*2, l_margin=6.35, r_margin=6.35)
        #
        description_p = cols.paragraph(text_align='L')
        pdf.set_font('dejavu-sans', size=12)
        description_lines = pdf.multi_cell(w=152.4, align='L', padding=(0, 11.175), text=description, dry_run=True, output="LINES")
        description_p.write('\n'.join(description_lines))
        cols.end_paragraph()
        #
        cols.new_column()
        #
        title_p = cols.paragraph(text_align='C')
        pdf.set_font('dejavu-sans', size=26)
        title_p.write(f"\n\n{title}")
        cols.end_paragraph()
        #
        author_p = cols.paragraph(text_align='C')
        pdf.set_font('dejavu-sans', size=14)
        author_p.write(f"\n{author}")
        cols.end_paragraph()
        #
        cols.render()
        #
        pdf.output(f"{folder}/{cover_pdf_fname}")
    return interior_pdf_fname, cover_pdf_fname, pages, pages >= 24 and pages <= 828

def get_books(run_folder, start, end):
    update_index_flag = True
    try:
        wb = openpyxl.load_workbook('Project Guttenberg.xlsx')
    except:
        wb = openpyxl.Workbook()
    try:
        del wb['Sheet']
    except:
        pass
    datestamp = datetime.now().strftime('%Y-%B-%d %H_%M')
    ws = wb.create_sheet(datestamp)
    ws.append(["Book ID", "Plain text URL", "Title", "Language", "Author", "Translator", "Illustrator", "Description", "Keywords", "BISAC codes", "Pages num", "PDF file name", "Cover PDF file name"])
    try:
        for i in range(start, end + 1):
            print(f'Processing index: {i}')
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
            book_txt = book_txt[book_content_start_index:book_content_end_index].replace('\r\n\r\n', '_____').replace('\r\n', '').replace('\n\n', '_____').replace('\n', '').replace('____', '\r\n\r\n').replace('____', '\n\n').replace('_', '')
            #
            description_query = f"Provide a 150 words description of the classic book {book_title}"
            if book_author:
                description_query += f" by Author and Writer {book_author}."
            if book_language:
                description_query += f" Write the review in this language: {book_language}"
            description_completion = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system",
                        "content": description_query
                    },
                ]
            )
            description = description_completion.choices[0].message.content
            #
            keywords_query = f'Give me 7 keywords separated by semicolons (only the keywords, no numbers nor introductory workds) for the classic book "{book_title}" by Author "{book_author}". Keywords must not be subjective claims about its quality, time-sensitive statments and must not include the word "book". Keywords must also not contain words included on the book the title, author nor contained on the following book description: {description}]'
            keywords_completion = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system",
                        "content": keywords_query
                    },
                ]
            )
            keywords = keywords_completion.choices[0].message.content
            #
            bisac_codes_query = f'Give me up to 3 BISAC codes (only the code, not its description and not numbered) for the book "{book_title}" by Author "{book_author}", for its correct classification'
            bisac_codes_completion = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system",
                        "content": bisac_codes_query
                    },
                ]
            )
            bisac_codes = bisac_codes_completion.choices[0].message.content
            #
            book_fname, cover_fname, pages_num, include_book_flag = generate_book_pdfs(run_folder, i, book_title, book_author, description, book_txt)
            #
            if include_book_flag:
                ws.append([i, book_txt_url, book_title, book_language, book_author, book_translator, book_illustrator, description, keywords, bisac_codes, pages_num, book_fname, cover_fname])
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
            update_last_index(end)

if __name__ == '__main__':
    # parse command line arguments
    parser = argparse.ArgumentParser(
        prog='guttenberg2.py', 
        usage='python3 %(prog)s [options]',
        description='Project Guttenberg books scrape script:',
        epilog="Script will create output folder named as datestamp, and also maintain last processed book index and Excel file with each run spreadsheet"
    )
    parser.add_argument('-s', '--start', type=int, dest='start', default=get_previous_last_index(), help='start index of the program')
    parser.add_argument('-e', '--end', type=int, dest='end', default=get_latest_published_book_index(), help='end index of the program')
    args = parser.parse_args()
    # create PDFs output folder
    run_folder = datetime.now().strftime('%Y-%B-%d')
    pathlib.Path(run_folder).mkdir(parents=True, exist_ok=True)
    #
    get_books(run_folder, args.start, args.end)
