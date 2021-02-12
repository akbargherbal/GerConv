import PySimpleGUI as sg
from itertools import chain
import os
from pyxpdf import Document
import subprocess


def compile_list(files_input='', directory=''):
    '''
    Compile a list of pdf files paths; for downstream processes.
    One shortcumming that needs to be addressed is the [;] delimiter.
    '''
    files_from_dir = []
    for root, dir_, files in os.walk(directory):
        for file in files:
            if file[-3:].lower() == 'pdf':
                files_from_dir.append(os.path.join(root, file))
    if files_input:
        files = files_input.split(';')
        return files + files_from_dir

    else:
        return files_from_dir
        
def copy_pdf_text(list_pdf, txt_file):
    '''
    Write pdf files text in [list_pdf] to a single TXT file [txt_file].
    '''
    list_failed_pdf_conversion = []
    for pdf_input in list_pdf:
        try:
            with open(txt_file, encoding='utf-8', mode='a+') as text_file:
                with open(pdf_input, mode='rb') as pdf_file:
                    doc = Document(pdf_file)
                    for page in doc:
                        text_file.write(f'{page.text()}\n')
        except:
            list_failed_pdf_conversion.append(pdf_input)
            print('The following pdf files couldn\'t be converted: ')
            for failed_file in list_failed_pdf_conversion:
                print(f'\t{failed_file}')



sg.theme('LightGrey2')     # Please always add color to your window

tab1_layout = [[sg.Text('Tab 1')]]


tab2_layout = [
                [sg.Text('Select Files', font='Roboto 12', size=(12,1), justification='right'),      sg.I(f'', key='-IN-F-TAB2-', size=(75,1)),                 sg.FilesBrowse(file_types=(("PDF Files", "*.pdf"),("PDF Files", "*.PDF")))],
                [sg.Text('Select Folder', font='Roboto 12', size=(12,1), justification='right'),     sg.I(f'{os.getcwd()}', key='-IN-D-TAB2-', size=(75,1)),    sg.FolderBrowse()],
                [sg.Text('Save File Name', font='Roboto 12', size=(12,1), justification='right'),     sg.I(f'{os.getcwd()}', key='-IN-SAV-TAB2-', size=(75,1)),              sg.FileSaveAs()],
                [sg.B("Start Conversion", key='-PDF_CONV-', font='Courier 12')],
                [sg.Output(size=(105,20))]
            ]

# The TabgGroup layout - it must contain only Tabs
tab_group_layout = [[sg.Tab('Dummy Tab',            tab1_layout, font='Courier 16', key='-TAB1-'),
                     sg.Tab('Convert pdf to txt',   tab2_layout, font='Courier 16', key='-TAB2-')
                    ]]

# The window layout - defines the entire window
layout = [[sg.TabGroup(tab_group_layout,
                       enable_events=True,
                       key='-TABGROUP-', font = 'Courier 16')]]


window = sg.Window('Convert to TXT format', layout, no_titlebar=False, size=(775,450))

while True:
    event, values = window.read()       # type: str, dict
    if event == sg.WIN_CLOSED:
        break

    lang_support = 'NOTE: ARABIC Language is NOT supported in PDF conversion!'
    
    files_, dir_, output_ = values['-IN-F-TAB2-'], values['-IN-D-TAB2-'], values['-IN-SAV-TAB2-']
    if (files_ == '') or (len(files_)>1):
        loop_counter = 0
        for file in files_.split(';'):
            if os.path.isfile(file) and (loop_counter<=0):
                file_dir_name = os.path.abspath(file)
                loop_counter += 1
            

    # print(loop_counter)
    if event == "-PDF_CONV-":
        sg.popup(lang_support, title='NOTE!', auto_close_duration=5, auto_close=True, font='Arial 14')
        list_pdf_files = compile_list(files_, dir_)
        list_pdf_files = [i.replace('\\', '/') for i in list_pdf_files]

        print(f'Number of PDF files to be converted = {len(list_pdf_files)}')

        list_check_pdf_files = [f for f in list_pdf_files if os.path.isfile(f)]
        if len(list_pdf_files) == len(list_check_pdf_files):
            copy_pdf_text(list_pdf_files, f'{output_}.txt')
            print('Converting pdf files was successful!')
            
            select_file = f'{output_}.txt'.replace('/', '\\')
            print('Output File is in the following folder:\n\t', select_file)
            subprocess.Popen(f'explorer /select,{select_file}')
            print('_'*50, '\n')

        # In case the select files path contains a special chracter like (;) which might corrupt splitting!
        else:
            warning_message = '''Files names must NOT contain special characters.
            Epecially characters like ; 
            Consider renaming those files'''


            sg.popup_scrolled(warning_message, title='WARNING! Invalid File Name!', size=(36,12), font='Arial 16')
            print(f'{file_dir_name}')
            subprocess.Popen(f'explorer /select,{file_dir_name}')

window.close()
