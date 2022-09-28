import re
import os
from os import path
from os import walk
from docx2pdf import convert
import pandas as pd
import numpy as np
from mailmerge import MailMerge
import PySimpleGUI as sg


def get_theme():
    """
    Get the theme to use for the program
    Value is in this program's user settings. If none set, then use PySimpleGUI's global default theme
    :return: The theme
    :rtype: str
    """
    # First get the current global theme for PySimpleGUI to use if none has been set for this program
    try:
        global_theme = sg.theme_global()
    except:
        global_theme = sg.theme()
    # Get theme from user settings for this program.  Use global theme if no entry found
    user_theme = sg.user_settings_get_entry('-theme-', '')
    if user_theme == '':
        user_theme = global_theme
    return user_theme


def combinar_correspondencia(contract, main_folder, output_folder_path, prefix, sufix, create_pdf, keep_word_file):
        print(contract['ID'])
        save_to_path = output_folder_path + str(contract['ID']) + "/"

        if not path.isdir(save_to_path):
            os.mkdir(save_to_path)

        filenames = next(walk(main_folder), (None, None, []))[2]  # [] if no file

        for index, file in enumerate(filenames):
            if "_template.docx" in file:
                print("   - ", str(file))
                with MailMerge(main_folder + file) as document:
                    exec(merge_code(document.get_merge_fields()))

                    document.write(
                        save_to_path + prefix + re.search(r"(?<=_)(.*)(?=_template.docx)", file)[0] + "_" +
                        contract["ID"] + sufix + ".docx")
                    print(re.search(r"(?<=_)(.*)(?=_template.docx)", file)[0], end="\r")

        if create_pdf:
            convert(output_folder_path + str(contract['ID']) + "/")

        if not keep_word_file:
            for index, file in enumerate(next(walk(output_folder_path + str(contract['ID']) + "/"), (None, None, []))[2]):
                if file.endswith(".docx"):
                    os.remove(os.path.join(output_folder_path + str(contract['ID']) + "/", file))


def merge_code(fields):
    length = len(fields)
    x = 0
    text = "document.merge("
    for field in fields:
        if field.find("money") > 0:
            text = text + str(
                field) + ' = str("$" + re.sub("(\d)(?=(\d{3})+(?!\d))", r"\\1.", "%d" % int(contract["' + str(
                field) + '"]))) + ",00"'
        elif field.find("number") > 0:
            text = text + str(field) + ' = str(re.sub("(\d)(?=(\d{3})+(?!\d))", r"\\1.", "%d" % int(contract["' + str(
                field) + '"])))'
        elif field.find("paragraph") > 0:
            text = text + str(field) + ' = str(contract["' + str(field) + '"]).replace(r"/\\n/g", "")'
        else:
            text = text + str(field) + '=str(contract["' + str(field) + '"])'
        if not x >= length - 1:
            text = text + ', '
        x += 1
    text = text + ")"
    return text


def main():

    # Define the window's contents

    # Define Combobox layouts
    layout_l = [
        [sg.Text("Prefijo de archivos creados")],
        [sg.Input('', enable_events=True, key='prefix')]
    ]

    layout_r = [
        [sg.Text("Sufijo de archivos creados")],
        [sg.Input('', enable_events=True, key='sufix')]
    ]

    # Define main window content
    layout = [
        [sg.FileBrowse('Buscar archivo Excel', file_types=(("MS Excel Files", "*.xlsx;*.xls"),),
                       key='-EXCELFILENAME-')],
        [sg.Col(layout_l, p=0), sg.Col(layout_r, p=0)],
        [sg.Checkbox('Crear PDFs', enable_events=True, key='create_pdf'),
         sg.Checkbox('Mantener documentos de word', enable_events=True, default=True, key='keep_word_file')],
        [sg.Text("AYUDA:", text_color="DARKRED", size=(20, 1), font=('Helvetica', 20),
                 justification='center')],
        [sg.Text("Siempre debe haber una columna ID la cual sera el nombre de cada carpeta.", text_color="BLACK")],
        [sg.Text(" - Si dentro del nombre de la columna se coloca la palabra 'money' esta tendra formato de moneda",
                 text_color="BLACK")],
        [sg.Text(" - Si dentro del nombre de la columna se coloca la palabra 'number' esta tendra formato de numero",
                 text_color="BLACK")],
        [sg.Text(
            " - Si dentro del nombre de la columna se coloca la palabra 'paragraph' esta tendra formato de parrafo sin saltos de linea",
            text_color="BLACK")],
        [sg.Text(
            "Recuerde: Solo seran combinados los documentos de word que en su nombre tengan la palabra '_template'")],
        [sg.T('   ' * 5), sg.Button('Combinar Correspondencia', size=(23, 1)),
         sg.T('   ' * 5),
         sg.Button('Salir', size=(23, 1))],
        [sg.ProgressBar(1000, orientation='h', size=(70, 40), key='progressbar', visible=False,
                        bar_color=('GREEN', 'WHITE'))],
        [sg.Text("", text_color="DARKBLUE", size=(20, 1), font=('Helvetica', 15), justification='center',
                 key='contratoTrabajoText', visible=False),
         sg.Text("", text_color="DARKGREEN", size=(20, 1), font=('Helvetica', 15), justification='right',
                 key='contratoActual', visible=False)]
    ]

    # Window - definition
    window = sg.Window('Combinacion de correspondencia (Word-Excel) a documentos separados',
                       layout)

    while True:
        # Display and interact with the Window
        event, values = window.read()  # Part 4 - Event loop or Window.read call
        if event in (sg.WIN_CLOSED, 'Salir'):
            # Finish up by removing from the screen
            break

        if not values['create_pdf'] and not values['keep_word_file']:
            window['keep_word_file'].Update(value=True)

        elif event == 'Combinar Correspondencia':
            window['contratoTrabajoText'].Update(visible=False)
            # If OK, then need to add the filename to the list of files and also set as the last used filename
            if not len(values["-EXCELFILENAME-"]) < 1:
                # Get excel folder
                excel_file = str(values["-EXCELFILENAME-"])
                main_folder = re.match(r'(?:([^<>:"\/\\|?*]*[^<>:"\/\\|?*.]\/|..\/))+', excel_file)[0]
                output_folder_path = re.match(r'(?:([^<>:"\/\\|?*]*[^<>:"\/\\|?*.]\/|..\/))+', excel_file)[
                                         0] + 'Output/'

                # Verif word file, excel file and folder (or create folder)
                if path.exists(excel_file):
                    print('contracts list file exist')
                    contracts = pd.read_excel(excel_file)
                    contracts = contracts.replace(np.nan, '', regex=True)
                else:
                    print('Error: contracts list file DO NOT exist')

                if not path.isdir(output_folder_path):
                    os.mkdir(output_folder_path)
                    print('Output folder created')

                window['progressbar'].Update(visible=True)

                for i, contract in contracts.iterrows():
                    window['contratoTrabajoText'].Update("Trabajando en: ", visible=True)
                    window['contratoActual'].Update(" " + contract['ID'], visible=True)
                    combinar_correspondencia(contract, main_folder, output_folder_path, values['prefix'],
                                             values['sufix'], values['create_pdf'], values['keep_word_file'])
                    window['progressbar'].UpdateBar(1000*(i/len(contracts)))
                    window['contratoTrabajoText'].Update(visible=False)
                    window['contratoActual'].Update(visible=False)

                # Do something with the information gathered
                window['contratoActual'].Update("")
                window['contratoTrabajoText'].Update("Finalizado con exito", visible=True)
                window['progressbar'].Update(visible=False)
                print('Successful execution!!!')

    window.close()  # Close Window


if __name__ == '__main__':
    main()
