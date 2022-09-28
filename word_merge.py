import os
import re
from os import path
from os import walk
import pandas as pd
import numpy as np
from docx2pdf import convert
from mailmerge import MailMerge


# ## Create all files related to de contract list (contracts) provided ## #
def create_process_and_contract_files():
    directory_path = "C:\\Users\\be47\\Documents\\IPCC\\2022\\08 agosto\\Combinacion otro si\\"
    excel_file = directory_path + 'Otrosi 2022.xlsx'
    #word_file = 'OtroSi_template.docx' ## _template.docx
    if path.exists(excel_file):
        print('contracts list file exist')
        contracts = pd.read_excel(excel_file)
        contracts = contracts.replace(np.nan, '', regex=True)
    else:
        print('Error: contracts list file DO NOT exist')

    for i, contract in contracts.iterrows():
        print(contract['(C) Número Del Contrato inicial'])
        folder_path = directory_path + 'Output'
        if not path.isdir(folder_path):
            os.mkdir(folder_path)

        filenames = next(walk(directory_path), (None, None, []))[2]  # [] if no file

        for index, word_file in enumerate(filenames):
            if "_template.docx" in word_file:
                with MailMerge(directory_path + word_file) as document:
                    document.merge(
                        Días_por_adicionar=str(contract["Días por adicionar"]),
                        F_Fecha_De_Terminación_Del_Contrato=str(contract["(F) Fecha De Terminación Del Contrato"]),
                        Valor_del_contrato_o_adicionar=str("$" + re.sub("(\d)(?=(\d{3})+(?!\d))", r"\1.",
                                                    "%d" % int(contract["Valor del contrato o adicionar"]))) + ",00",
                        D_No_Disponibilidad_Presupuestal=str(contract["(D) No. Disponibilidad Presupuestal"]),
                        C_Nombre_Completo_Del_Contratista=str(contract["(C) Nombre Completo Del Contratista"]),
                        Plazo=str(contract["Plazo"]),
                        C_Número_Del_Contrato_inicial=str(contract["(C) Número Del Contrato inicial"]),
                        C_Objeto_Contractual=str(contract["(C) Objeto Contractual"]).replace(r'/\n/g', "")
                    )

                    document.write(folder_path + "\\" + "Otro_si_" +
                                   contract["(C) Número Del Contrato inicial"] + ".docx")
                    print(folder_path + "\\" + "Otro_si_" +
                                   contract["(C) Número Del Contrato inicial"] + ".docx", end="\r")
