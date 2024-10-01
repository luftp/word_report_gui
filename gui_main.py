from gui_classes import GUI_relatorio


template_path = r"template.docx"

path_saida = r'docteste.docx'

path_figuras = r"figuras\Capturar.PNG"

dict_estilos = {1: 'EPC_Titulo1',
                2: 'EPC_Titulo2',
                3: 'EPC_Titulo3',
                4: 'EPC_Titulo4',
                5: 'EPC_Titulo5',
                6: 'EPC_Titulo6'}

gui1 = GUI_relatorio(dict_estilos, template_path, path_saida, path_figuras)

