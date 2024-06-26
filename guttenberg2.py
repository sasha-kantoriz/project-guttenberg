import requests
from datetime import datetime
import pathlib
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

def generate_book_pdfs(folder, _id, title, author, description, preface, contents, text, cover_only=False, word_only=False):
    interior_pdf_fname, cover_pdf_fname, front_cover_pdf_fname = f"{_id}_paperback_interior.pdf", f"{_id}_paperback_cover.pdf", f"{_id}_paperback_front_cover.pdf"
    pdf = PDF(format=(152.4, 228.6))
    pdf.add_font("dejavu-sans", style="", fname="assets/DejaVuSans.ttf")
    # TITLE
    pdf.add_page()
    pdf.set_font("dejavu-sans", size=24)
    lines_num = len(pdf.multi_cell(w=0, align='C', padding=(0, 8), text=f"{title}\n\n{author}", dry_run=True, output="LINES"))
    if lines_num >= 3:
        padding_top = (228.6 - 24 * (lines_num - 1)) / 2
    else:
        padding_top = (228.6 - 24 * (lines_num)) / 2
    pdf.multi_cell(w=0, align='C', padding=(padding_top, 8, 0), text=f"{title}\n\n{author}")
    if preface or contents:
        # PREFACE
        pdf.add_page()
        pdf.set_font("dejavu-sans", size=10)
        pdf.multi_cell(w=0, align='L', padding=8, text=preface)
        # CONTENTS
        pdf.add_page()
        pdf.set_font("dejavu-sans", size=10)
        pdf.multi_cell(w=0, align='C', padding=8, text=contents)
    # TEXT
    pdf.add_page()
    pdf.set_font("dejavu-sans", size=10)
    pdf.multi_cell(w=0, h=4.6, align='J', padding=8, text=text)
    #
    pages = pdf.page_no()
    if pages >= 24 and pages <= 828 and not cover_only and not word_only:
        pdf.output(f"{folder}/pdf/{interior_pdf_fname}")
    # COVERS
    if pages >= 24 and pages <= 828 and not word_only:
        # FRONT COVER
        cover_width, cover_height = 152.4 + 3.175, 234.95
        pdf = fpdf.FPDF(format=(cover_width, cover_height))
        pdf.add_font('dejavu-sans', style="", fname="assets/DejaVuSans.ttf")
        pdf.add_page()
        pdf.set_fill_color(r=250, g=249, b=222)
        pdf.rect(h=pdf.h, w=pdf.w, x=0, y=0, style="DF")
        pdf.set_font('dejavu-sans', size=18)
        text_h = pdf.multi_cell(w=0, align='C', padding=6.35, text=f"\n\n{title}\n\n* * *\n\n{author}", dry_run=True, output="HEIGHT")
        pdf.multi_cell(w=0, align='C', padding=6.35, text=f"\n\n{title}\n\n* * *\n\n{author}")
        # COVER IMAGE
        include_cover_img = (text_h + 8) < (134.95 / 2)
        #
        if include_cover_img:
            try:
                cover_img = f'{folder}/imgs/{_id}.png'
                prompt = f"Generate an image to be used as a part of a classic book cover, without any text letters or words on the image, reflecting the following description: {description}. The image needs to be without words, letters or any text and not contain the book with its cover"
                img_url = client.images.generate(model='dall-e-3', prompt=prompt, n=1, quality="standard").data[0].url
                response = requests.get(img_url)
                with open(cover_img, 'wb') as img:
                    img.write(response.content)
                pdf.image(cover_img, x=(152.4 - 100 + 6.35) / 2,
                          y=(234.95 - 40) / 2, w=100, h=100)
            except:
                pass
        pdf.output(f"{folder}/front_cover/{front_cover_pdf_fname}")
        # Full cover
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
        pdf.set_font('dejavu-sans', size=28)
        title_h = pdf.multi_cell(w=0, align='C', padding=(0, 8), text=f"\n\n{title}", dry_run=True, output="HEIGHT")
        title_p.write(f"\n\n{title}")
        cols.end_paragraph()
        #
        separator_text = "\n* * *"
        separator_p = cols.paragraph(text_align='C')
        pdf.set_font('dejavu-sans', size=16)
        separator_h = pdf.multi_cell(w=0, align='C', padding=(0, 8), text=separator_text, dry_run=True, output="HEIGHT")
        separator_p.write(separator_text)
        cols.end_paragraph()
        #
        author_p = cols.paragraph(text_align='C')
        pdf.set_font('dejavu-sans', size=16)
        author_h = pdf.multi_cell(w=0, align='C', padding=(0, 8), text=f"\n{author}", dry_run=True, output="HEIGHT")
        author_p.write(f"\n{author}")
        cols.end_paragraph()
        #
        if include_cover_img:
            try:
                pdf.image(cover_img, x=(152.4 + pages * 0.05720 + 3.175) + (152.4 - 100 - 6.35) / 2, y=(234.95 - 40) / 2, w=100, h=100)
            except:
                pass
        #
        cols.render()
        #
        pdf.output(f"{folder}/cover/{cover_pdf_fname}")
    return interior_pdf_fname, cover_pdf_fname, front_cover_pdf_fname, pages, pages >= 24 and pages <= 828

def generate_book_docx(folder, _id, title, author, description, preface, contents, text):
    doc = docx.Document("assets/template.docx")
    currentYear, currentMonth = datetime.now().year, datetime.now().month
    title_paragraph = doc.add_paragraph()
    title_paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_paragraph.add_run(f"{title}\n\n{author}")
    title_font = title_run.font
    title_font.name = 'Verdana'
    title_font.size = docx.shared.Pt(24)
    doc.add_page_break()
    if preface:
        preface_paragraph = doc.add_paragraph()
        preface_run = preface_paragraph.add_run(preface)
        preface_font = preface_run.font
        preface_font.name = 'Verdana'
        preface_font.size = docx.shared.Pt(10)
        doc.add_page_break()
    if contents:
        contents_paragraph = doc.add_paragraph()
        contents_run = contents_paragraph.add_run(contents)
        contents_font = contents_run.font
        contents_font.name = 'Verdana'
        contents_font.size = docx.shared.Pt(10)
        doc.add_page_break()
    text_paragraph = doc.add_paragraph()
    text_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
    text_run = text_paragraph.add_run(text)
    text_font = text_run.font
    text_font.name = 'Verdana'
    text_font.size = docx.shared.Pt(10)
    doc.save(f"{folder}/word/{_id}_paperback_interior.docx")


def get_books(run_folder, start, end, cover_only=False, word_only=False, indexes=None):
    update_index_flag = True
    datestamp = datetime.now().strftime('%Y-%B-%d %H_%M')
    if not (cover_only or word_only):
        try:
            wb = openpyxl.load_workbook('Project Guttenberg.xlsx')
        except:
            wb = openpyxl.Workbook()
        try:
            del wb['Sheet']
        except:
            pass
        ws = wb.create_sheet(datestamp)
        ws.append(["Book ID", "Plain text URL", "Title", "Published Year", "Language", "Author", "Author Year of Death", "Translator", "Illustrator", "Description", "Keywords", "BISAC codes", "Pages num", "PDF file name", "Cover PDF file name", "Front cover PDF file name"])
    try:
        sequence = indexes if indexes else range(start, end + 1)
        for i in sequence:
            print(f'Processing index: {i}')
            sleep(randint(1, 3))
            book_url = f'https://www.gutenberg.org/ebooks/{i}'
            book_txt_url = f'{book_url}.txt.utf-8'
            book_txt = requests.get(book_txt_url, timeout=60, headers={'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 6.1; zh-CN) AppleWebKit/533+ (KHTML, like Gecko)'}).content.decode('utf-8')
            #
            book_author = re.search(r"Author: (.*)\n", book_txt)
            book_author = book_author.groups()[0] if book_author else ""
            book_language = re.search(r"Language: (.*)\n", book_txt, re.IGNORECASE)
            book_language = book_language.groups()[0] if book_language else ""
            book_translator = re.search(r"Translator: (.*)\n", book_txt, re.IGNORECASE)
            book_translator = book_translator.groups()[0] if book_translator else ""
            book_illustrator = re.search(r"Illustrator: (.*)\n", book_txt, re.IGNORECASE)
            book_illustrator = book_illustrator.groups()[0] if book_illustrator else ""
            book_title = re.search(r"Title: (.*)\n", book_txt)
            book_title = book_title.groups()[0] if book_title else ""
            book_content_start_index = re.search(r"\*\*\* START OF THE PROJECT GUTENBERG .* \*\*\*", book_txt, re.IGNORECASE)
            book_content_start_index = book_content_start_index.end() if book_content_start_index else 0
            book_content_end_index = re.search(r"\*\*\* END OF THE PROJECT GUTENBERG .* \*\*\*", book_txt, re.IGNORECASE)
            book_content_end_index = book_content_end_index.start() if book_content_end_index else -1
            book_txt = book_txt[book_content_start_index:book_content_end_index]
            #
            if "hungarian" in book_language.lower() or "romanian" in book_language.lower() or "esperanto" in book_language.lower() or "latin" in book_language.lower() or not book_author:
                continue
            #
            book_txt = re.sub(r'\[Illustration[^\]]*\]', '', book_txt)
            book_txt = book_txt.replace('\r\n', '\n')
            #
            contents_search = re.search(r"(content|contents|chapters)(:)?\n\n", book_txt, re.IGNORECASE)
            if contents_search and not re.search(r"(content|contents|chapters)(:)?(\n)+(\s)*of", book_txt[:contents_search.start() + 100], re.IGNORECASE):
                contents_start_index = contents_search.start()
                contents_end_index = contents_start_index + len(contents_search.groups()[0]) + 5 + book_txt[contents_start_index + len(contents_search.groups()[0]) + 5:].find('\n\n\n')
            else:
                contents_end_index = contents_start_index = re.search(r"\n\n\n", book_txt).start()
            #
            book_preface = book_txt[:contents_start_index]
            book_contents = book_txt[contents_start_index:contents_end_index]
            book_txt = book_txt[contents_end_index:]
            #
            book_preface = book_preface.replace('\n\n\n\n', '\n\n').replace('_', '').replace('  ', ' ').replace('--', '-').replace('\n\n', '_____').replace('\n', ' ').replace('_____', '\n\n')
            book_contents = book_contents.replace('\n\n\n\n', '\n\n').replace('_', '').replace('  ', ' ').replace('--', '-').replace('\n\n', '\n')
            book_txt = book_txt.replace('\n\n\n\n', '\n\n').replace('_', '').replace('  ', ' ').replace('--', '-').replace('\n\n', '_____').replace('\n', ' ').replace('_____', '\n\n')
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
            keywords_query = f'Give me 7 keywords separated by semicolons (only the keywords, no numbers nor introductory words) for the classic book "{book_title}" by Author "{book_author}". Keywords must not be subjective claims about its quality, time-sensitive statments and must not include the word "book". Keywords must also not contain words included on the book the title, author nor contained on the following book description: {description}'
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
            bisac_codes_query = f'Give me up to 3 BISAC codes separated by semicolons (only the code, not its description and not numbered) for the book "{book_title}" by Author "{book_author}" with description "{description}", for its correct classification. Output format example would be: FIC019000; FIC031010; FIC014000'
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
            published_year_query = f'Please, tell me the year the book {book_title} by {book_author} was published. Provide only the date in the format YYYY.'
            published_year_completion = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system",
                        "content": published_year_query
                    },
                ]
            )
            published_year = published_year_completion.choices[0].message.content
            #
            author_year_of_death_query = f'Please, tell me the year of death of {book_author}, the author of the book {book_title}. Provide only the date in the format YYYY. If the author is still alive, please, provide the "XXXX".'
            author_year_of_death_completion = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system",
                        "content": author_year_of_death_query
                    },
                ]
            )
            author_year_of_death = author_year_of_death_completion.choices[0].message.content
            #
            book_fname, cover_fname, front_cover_fname, pages_num, include_book_flag = generate_book_pdfs(run_folder, i, book_title, book_author, description, book_preface, book_contents, book_txt, cover_only, word_only)
            #
            if (not cover_only or word_only) and (pages_num >= 24 and pages_num <= 828):
                generate_book_docx(run_folder, i, book_title, book_author, description, book_preface, book_contents, book_txt)
            #
            if include_book_flag and not (cover_only or word_only):
                ws.append([i, book_txt_url, book_title, published_year, book_language, book_author, author_year_of_death, book_translator, book_illustrator,
                           description, keywords, bisac_codes, pages_num, book_fname, cover_fname, front_cover_fname])
    except KeyboardInterrupt:
        update_index_flag = False
    except Exception as e:
        print(e)
        update_index_flag = False
        update_last_index(i)
    finally:
        if not (word_only or cover_only):
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
    parser.add_argument('--indexes', dest='indexes', help='books indexes to process, comma separated')
    parser.add_argument('-s', '--start', type=int, dest='start', default=get_previous_last_index(), help='start index of the program')
    parser.add_argument('-e', '--end', type=int, dest='end', default=get_latest_published_book_index(), help='end index of the program')
    parser.add_argument('--word', action='store_true', help='generate Word documents')
    parser.add_argument('--cover', action='store_true', help='generate PDF covers')
    args = parser.parse_args()
    # create PDFs output folder
    run_folder = datetime.now().strftime('%Y-%B')
    pathlib.Path(f"{run_folder}/imgs").mkdir(parents=True, exist_ok=True)
    pathlib.Path(f"{run_folder}/cover").mkdir(parents=True, exist_ok=True)
    pathlib.Path(f"{run_folder}/front_cover").mkdir(parents=True, exist_ok=True)
    pathlib.Path(f"{run_folder}/word").mkdir(parents=True, exist_ok=True)
    pathlib.Path(f"{run_folder}/pdf").mkdir(parents=True, exist_ok=True)
    #
    get_books(run_folder, args.start, args.end, args.cover, args.word, args.indexes.split(',') if args.indexes else None)
