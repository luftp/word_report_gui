import re
from word_report import Word_report

def parser_paragrafo(texto):

    list_comandos = re.findall( '#([^#]*)#', texto)

    return_comandos = []

    for comando in list_comandos:

        comando_parser = comando.replace(' ','')

        if comando_parser.startswith('fig.'):

            lista_seq =  comando.split('.')

            id_comando = lista_seq[0]

            id_fig =  lista_seq[1]

            legend_fig = lista_seq[2]

            return_comandos.append([id_comando, id_fig, legend_fig])

            texto = texto.replace('#'+comando+'#', '')


        if comando_parser.startswith('ref.'):
            
            lista_seq =  comando.split('.')

            id_comando = lista_seq[0]

            id_fig =  lista_seq[1]
        
            return_comandos.append([id_comando, id_fig])

            texto = texto.replace('#'+comando+'#', '')


    return return_comandos, texto



def compilar_relatorio(lista_itens, template, dict_estilos, path_saida_relatorio):

        if len(lista_itens) > 0:

            mc = Word_report(template)

            for item in lista_itens:

                if item.get_tipo() == 'titulo':
            
                    lista_values = list(dict_estilos.values())
                    lista_keys = list(dict_estilos.keys())

                    mc.add_item(item.get_texto(), int(lista_keys[lista_values.index(item.get_estilo())]))
                
                elif item.get_tipo() == 'paragrafo':

                    lista_comandos, texto_parser = parser_paragrafo(item.get_texto())

                    mc.add_paragrafo(texto_parser)

                    for comando in lista_comandos:

                        if comando[0] == 'fig':

                            mc.add_figura(r"C:\Users\lftavares\OneDrive - DF+ ENGENHARIA\Documents\PYTHON\ENGENHARIA\MEMORIA_CALCULO\calculation_report\nicegui\figuras\Capturar.PNG", comando[2], bookmark=comando[1])

                        if comando[0] == 'ref':

                            mc.add_ref_to_figure('', comando[1])



            mc.save_file(path_saida_relatorio)

