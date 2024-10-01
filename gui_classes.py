from nicegui import ui
from gui_compilador import compilar_relatorio
import json

   

        
class GUI_relatorio:

    '''Classe para gerar uma interface gráfica para compilação de relatórios

    A classe utiliza a biblioteca nicegui para criar uma interface gráfica em que o usuário pode construir a estrutura de um relatório, permitindo compilar em um
    documento Word.

    Attibutes:
        dict_estilos: dicionário com os nomes de estilos que serão utilizados no Word
        template: path para documento Word com os estilos e template para o relatório final
        path_saida_relatorio: path para saida do relatório compilado
        path_figuras: caminho para a pasta com as figuras que serão utilizadas no relatório compilado
        lista_itens: lista com os elementos da GUI para a compilação do relatório
        container_time_line: "timeline" da GUI com todos os elementos utilizados para compilar o relatório
        header: header com todas as funções da GUI
    '''


    def __init__(self, dict_estilos, template, path_saida_relatorio, path_figuras):

        self.dict_estilos = dict_estilos
        self.template = template
        self.path_saida_relatorio = path_saida_relatorio
        self.path_figuras = path_figuras

        self.lista_itens = []

        self.container_time_line = ui.column().classes("w-9/12 self-center")

        self.header = ui.header()


        with self.header:

            ui.button('Adicionar um titulo', on_click=lambda: self.add_titulo())

            ui.button('Adicionar um paragrafo', on_click=lambda: self.add_paragrafo())

            ui.button('Gerar relatório', on_click=lambda: self.gerar_relatorio())

            ui.button('Limpar', on_click= lambda: self.limpar_itens())

            ui.button('Salvar', on_click=lambda: ui.download(self.salvar_estrutura()))

            ui.upload(on_upload= lambda e: self.upload_estrutura(e), auto_upload=True).props('accept=.json label=Carregar')

        '''
        Comandos para inicializar a GUI utilizando dark mode
        '''

        dark = ui.dark_mode()

        dark.enable()

        ui.run()
            

    def limpar_itens(self):

        '''Função que limpa todos os elementos da GUI
        '''

        self.lista_itens = self.refresh_lista_itens()

        for item in self.lista_itens:

            item.clear_item()

    def refresh_lista_itens(self):

        '''Função para atualizar a lista de elementos = Verifica se o state de algum item está como None e o remove da lista se verdadeiro.
        '''

        return [item for item in self.lista_itens if item.state != None]


    def gerar_relatorio(self):

        
        '''Função compilar o relatório. Utiliza pacotes externos
        '''

        lista_itens = self.refresh_lista_itens()

        compilar_relatorio(lista_itens, self.template, self.dict_estilos, self.path_saida_relatorio)


    def add_titulo(self):

        '''Função que adiciona um elemento do tipo título
        '''

        with self.container_time_line:

            self.lista_itens.append(Elemento('titulo','Texto Título', 'Entrar com Título', list(self.dict_estilos.values())))

    def add_paragrafo(self):

        
        '''Função que adiciona um elemento do tipo parágrafo
        '''

        with self.container_time_line:

            self.lista_itens.append(Elemento('paragrafo','Texto Parágrafo', 'Escrever texto de um parágrafo', list(self.dict_estilos.keys())))



    def upload_estrutura(self, e):

        '''Função para carregar na GUI uma estrutura salva anteriormente. Deve ser um arquivo estruturado especificamente para a GUI, com tipo JSON
        '''

        data = e.content.read().decode('utf-8')

        data = json.loads(data)

        self.limpar_itens()

        self.dict_estilos = data['dict_estilo']

        self.template = data['template']
        
        self.path_saida_relatorio = data['path_saida']

        self.path_figuras = data['path_figuras']

        itens_adicionar = data['dict_itens']


        for item in itens_adicionar:

            if item[0] == 'titulo':

                self.add_titulo()

                self.lista_itens[-1].set_estilo(item[1])

                self.lista_itens[-1].set_texto(item[2])

            elif item[0] == 'paragrafo':

                self.add_paragrafo()

                self.lista_itens[-1].set_estilo(item[1])

                self.lista_itens[-1].set_texto(item[2])


    def salvar_estrutura(self):

        '''Função salvar um arquivo tipo JSON com a estrutura atual carregada na GUI.
        '''

        self.refresh_lista_itens()

        dict_estrutura = dict()

        dict_estrutura['dict_estilo'] = self.dict_estilos

        dict_estrutura['template'] = self.template

        dict_estrutura['path_saida'] = self.path_saida_relatorio

        dict_estrutura['path_figuras'] = self.path_figuras

        dict_estrutura['dict_itens'] = []

        for item in self.lista_itens:

            dict_estrutura['dict_itens'].append([item.get_tipo(), item.get_estilo(), item.get_texto()])

        with open('data.json', 'w') as f:
        
            json.dump(dict_estrutura, f, indent=4)

        return 'data.json'


                        




    


class Elemento:

    def __init__(self, tipo, label, placeholder, lista_estilo):

        self.tipo = tipo
        self.label = label
        self.placeholder = placeholder
        self.lista_estilo = lista_estilo
        self.state = ''

        if self.tipo == 'titulo':

            self.container = ui.card().classes("self-start")

            with self.container:

                with ui.row():

                    self.texto = ui.input(label=self.label, placeholder=self.placeholder)

                    self.estilo_select =ui.select(self.lista_estilo , value=self.lista_estilo[0], on_change= lambda e: self.set_estilo(e.value))

                    self.set_estilo(self.lista_estilo[0])

                    self.button = ui.button('Remover', on_click=self.clear_item)
    
        elif self.tipo == 'paragrafo':

            self.container = ui.card().classes("w-full")

            with self.container:

                with ui.expansion(caption='Parágrafo', value=True).classes('w-full'):

                    self.texto = ui.textarea(label=self.label, placeholder=self.placeholder).classes('w-full')

                    self.estilo = ''

                    self.button = self.button = ui.button('Remover', on_click=self.clear_item)


    def get_tipo(self):

        return self.tipo
    
    def set_texto(self, value):

        self.texto.value = value
        
    def get_texto(self):

        return self.texto.value
    
    def set_estilo(self, estilo):

        self.estilo = estilo

    def get_estilo(self):

        return self.estilo

    def clear_item(self):

        self.container.delete()

        self.state = None
