import PySimpleGUI as sg
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import os

def create_output_folder(output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

def extract_text_from_image(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img)
    return text

def capture_image(pdf_path, page_num, bbox, output_folder):
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    image_name = f'page_{page_num + 1}.png'
    image_path = os.path.join(output_folder, image_name)
    img.save(image_path)
    doc.close()

    return image_path

def search_word_in_pdf(word, pdf_path, output_folder):
    found_info = []

    doc = fitz.open(pdf_path)

    for page_num in range(doc.page_count):
        page = doc[page_num]
        text = page.get_text("text")
        for paragraph in text.split('\n'):
            if word.lower() in paragraph.lower():
                bbox = page.search_for(word)
                found_info.append((page_num, bbox))

    doc.close()

    create_output_folder(output_folder)

    captured_images = []
    for page_num, bbox in found_info:
        image_path = capture_image(pdf_path, page_num, bbox, output_folder)
        captured_images.append(image_path)

    return captured_images

def main():
    layout = [
        [sg.Text('Enter a word:'), sg.InputText(key='word')],
        [sg.Text('Select PDF directory:'), sg.InputText(key='pdf_dir'), sg.FolderBrowse()],
        [sg.Text('Select Output Folder:'), sg.InputText(key='output_folder', enable_events=True), sg.FolderBrowse()],
        [sg.Button('Search'), sg.Button('Exit')],
        [sg.Image(filename='', key='image')],
        [sg.Multiline('', size=(60, 5), key='output', autoscroll=True)]
    ]

    window = sg.Window('PDF Word Search', layout)

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            break

        if event == 'Search':
            word = values['word']
            pdf_dir = values['pdf_dir']
            output_folder = values['output_folder']

            if not word or not pdf_dir or not output_folder:
                sg.popup_error('Please enter a word, select a PDF directory, and choose an output folder.')
                continue

            pytesseract.pytesseract.tesseract_cmd = pytesseract.get_tesseract_version()[0]  # Automatically determine Tesseract installation location

            pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]

            if not pdf_files:
                sg.popup_error('No PDF files found in the selected directory.')
                continue

            for pdf_file in pdf_files:
                pdf_path = os.path.join(pdf_dir, pdf_file)
                captured_images = search_word_in_pdf(word, pdf_path, output_folder)

                if captured_images:
                    window['output'].update(f'Word found in {len(captured_images)} pages.')
                    window['image'].update(filename=captured_images[0])
                else:
                    window['output'].update('Word not found in any PDFs.')

    window.close()

if __name__ == '__main__':
    main()
