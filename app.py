import PySimpleGUI as sg
import pandas as pd
from funcItens import Injecao, Sopro
import base64


with open("favicon.png", "rb") as img_file:
    favicon = base64.b64encode(img_file.read())

#GuiTheme
sg.theme("SystemDefault")
# faviocn

# Main view
def start_program():
    df = pd.read_csv("Data/product_database.csv",sep="|")
    df = df.dropna(axis = 0, how = 'all')
    pigmento = pd.read_csv("Data/pigmento.csv", sep=";")
    pigmento = pigmento.dropna()
    #lista de emails para enviar
    email_list = pd.read_csv("Data/Email.txt")
    
    #Gui da injeção
    injecao = [
        [sg.Text('*Máquina:'), sg.Input('', size=(10, 30), key='-Maquina-'),sg.Text('*Produto: '), sg.Combo(df.query("Setor == 'Injecao'").Produto.to_list(), size=(15), key='-Itens-')],
        [sg.Text('*Quant.  :'), sg.Input('', size=(20), key='-Quantidade-')],
        [sg.Text('*Cliente:  '), sg.Input('', size=20, key='-Cliente-')],
        [sg.Text('*Cor:       '), sg.Input('', size=(20, 50), key='-Cor-')],
        [sg.Text('COD:      '), sg.Combo(pigmento.cor.to_list(), size=(20,50), key='-COD-')],
        [sg.Text('OBS:      '), sg.Input('', size=(30, 2), key='-OBS-')],
        [sg.Radio(group_id=2, text='Rosca',default=True,key='-bocal_rsc-'),sg.Radio(group_id=2, text='Batoque',key='-bocal_bt-'),sg.Radio(group_id=2, text='Sem Bocal',key='-sem-')],
        [sg.Text('Embalagem:')], [sg.Radio(group_id='1', text='Caixa', default=True, key='-Emb_caixa-'),
                                sg.Radio(group_id='1', text='Bag', key='-Emb_bag-'),
                                sg.Radio(group_id='1', text='Saco', key='-Emb_saco-')],
        
        [sg.Text('Ações:')], [sg.Checkbox('Imprimir', default=True, key='-Imprimir-'),
                            sg.Checkbox('Enviar Email', default=True, key='-Send_email-')],
        [sg.Button('Enviar', button_color=(sg.YELLOWS[0], sg.GREENS[0]),expand_x=True),sg.Button('Sair', button_color=(sg.YELLOWS[0], 'Red'),expand_x=True)]

    ]
    #Gui do Sopro
    sopro = [
        [sg.Text('*Máquina:    '), sg.Input('', size=(10, 30), key='-sMaquina-')],
        [sg.Text('*Produto:     '), sg.Combo(df.query("Setor == 'Sopro'").Produto.to_list(), size=(20), key='-sItens-')],
        [sg.Text('*Quantidade:'), sg.Input('', size=(20), key='-sQuantidade-')],
        [sg.Text('*Cliente:      '), sg.Input('', size=20, key='-sCliente-')],
        [sg.Radio(group_id=3, text='Rosca',default=True,key='-Sbocal_rsc-'),sg.Radio(group_id=3, text='Batoque',key='-Sbocal_bt-')],
        [sg.Text('*Cor:           '), sg.Input('', size=(20, 50), key='-sCor-')],
        [sg.Text(' OBS:         '), sg.Input('', size=(30, 2), key='-sOBS-')],
        [sg.Text('Ações:')], [sg.Checkbox('Imprimir', default=True, key='-sImprimir-'),
                            sg.Checkbox('Enviar Email', default=True, key='-sSend_email-')],
        [sg.Button('Enviar', button_color=(sg.YELLOWS[0], sg.GREENS[0]),expand_x=True,k='-Enviar-'),sg.Button('Sair', button_color=(sg.YELLOWS[0], 'Red'),expand_x=True,k='-Sair-')]
    ]
    #Caixa que adicona tudo
    layout = [
        [sg.MenubarCustom([['Email', ['Login','lista trans.']], ['Produto', ['Produtos','pigmentos']]],
                          bar_background_color="blue",
                          bar_text_color="White",
                          p=0)],
        [sg.TabGroup(
            [[
                sg.Tab('Injeção',injecao),
                sg.Tab('Sopro',sopro) ]]),
                sg.Output(size=(30,17),key='-Output-')
            ],

            [sg.Text('Criado por Levi Renato', font='italic')],

    ]

    # Event Loop para abrir a janela e receber os valores e eventos
    main = sg.Window('BetelSystem', layout, icon="favicon.ico")

    while True:
        event, values = main.read()  # type: ignore
        # Eventos
        if event == sg.WIN_CLOSED or event == 'Sair' or event == '-Sair-':
            break
        if event == "Login":
            main.close()
            editar_email()
            break
        elif event == "lista trans.":
            main.close()
            editar_lista_email()
            break
        elif event == "Produtos":
            main.close()
            editar_produto()
            break
        elif event == "pigmentos":
            main.close()
            editar_pigmento()
            break
        # Variavéis
        numero_maquina = values['-Maquina-']
        quantidade = values['-Quantidade-']
        cliente = values['-Cliente-']
        produto = values['-Itens-']
        cor = values['-Cor-']
        if values['-COD-'] == "":
            Cod = ""
        else:Cod = pigmento.query("cor == '{}'".format(values['-COD-'])).codigo.to_string(index=False)
        
        obs = values['-OBS-']
        # volume da embalagem
        vol = ''
        if values['-Emb_caixa-']:
            vol = 'Cx'
        elif values['-Emb_bag-']:
            vol = 'Bag'
        elif values['-Emb_saco-']:
            vol = 'Saco'
        else:
            print('Error')
            
        # Bocal
        bocal = ''
        if values ['-bocal_rsc-']:
            bocal = "Rosca"
        elif values['-sem-']:
            bocal = ""
        else: bocal = "Batoque"
        
        obs = values['-OBS-']
        # Ao apertar enviar vai checar se os seguintes campos estão vazios:
        if event == 'Enviar':
            if not numero_maquina and numero_maquina == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            elif not produto and produto == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            elif not cor and cor == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            # quantidade
            elif not quantidade and quantidade == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            # cliente
            elif not cliente and cliente == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            else:
                #Após checagem vai colocar as variaveis na função para substituir na planilha
                try:
                     resposta_injecao = Injecao(df,numero_maquina,produto,cor,quantidade,bocal,cliente,Cod,obs,vol)
                except:
                    sg.popup_error('Digite uma quantidade válida!')

                if values['-Imprimir-']:
                    resposta_injecao.imprimir()
                    print('Arquivo enviado p/ impressão')
                else:
                    pass
                if values['-Send_email-']:
                    for i in email_list.email.dropna():
                        try:
                            resposta_injecao.send_email(i)
                            print(f'Email enviado p/ {i}')
                        except: print('Falha ao enviar email')
                else:
                    pass
                print("Arquivo {} criado".format(resposta_injecao.name))
                print('==========================')
                sg.popup('Programação Feita!')

    ### SOPRO ###
        snumero_maquina = values['-sMaquina-']
        squantidade = values['-sQuantidade-']
        scliente = values['-sCliente-']
        sproduto = values['-sItens-']
        scor = values['-sCor-']
        sobs = values['-sOBS-']
        # Bocal
        sbocal = ''
        if values ['-Sbocal_rsc-']:
            sbocal = "Rosca"
        else: sbocal = "Batoque"
        
        if event == '-Enviar-':
            if not snumero_maquina and snumero_maquina == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            elif not sproduto and sproduto == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            elif not scor and scor == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            # quantidade
            elif not squantidade and squantidade == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            # cliente
            elif not scliente and scliente == '':
                sg.popup_error('Não deixe os campos com * vazio!')
            else:
                try:
                   resposta_sopro = Sopro(df,snumero_maquina,sproduto,scor,squantidade,sbocal,scliente,sobs)
                except:
                    sg.popup_error('Digite uma quantidade válida!')

                if values['-sImprimir-']:
                    resposta_sopro.imprimir()
                    print('Arquivo enviado p/ impressão')
                else:
                    pass
                if values['-sSend_email-']:
                    for i in email_list.email.dropna():
                        try:
                            resposta_sopro.send_email(i)
                            print(f'Email enviado p/ {i}')
                        except:
                            print('Falha ao enviar email')
                else:
                    pass
                print("Arquivo {} criado".format(resposta_sopro.name))
                print('==========================')
                sg.popup('Programação Feita!')
    main.close_destroys_window()

# edit email to sender
def editar_email():
    fn = pd.read_csv("Data/UserEmail.config", sep=";")
 
    layout_login = [
                [sg.Text('Email')],
                [sg.Input(fn.email.to_string(index=False),key='-email-')],
                [sg.Text('Senha')],
                [sg.Input(fn.senha.to_string(index=False),key='-senha-',password_char="*")],
                [sg.Button('Salvar')]  ]
    
    
    window = sg.Window('Editar Email', layout_login,use_custom_titlebar=True, titlebar_icon=favicon)
    while True:             # Event Loop
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            window.close()
            break
        if event == 'Salvar':
            try:
                fn["email"] = fn.email.replace(fn.email.to_string(index=False), values['-email-'])
                fn["senha"] = fn.senha.replace(fn.senha.to_string(index=False), values['-senha-'])
                fn.to_csv("Data/USerEmail.config", sep=";", index=False)
                sg.popup_quick_message("Salvo!")
            except: sg.popup_error("Ops! Algo deu errado")
    start_program()

# edit email list
def editar_lista_email():
    email_list = pd.read_csv("Data/Email.txt")
    layout_lista = [
        [sg.Text("Lista de transmissão", font=('Helvetica', 13), justification="center")],
        [sg.Text("Email 1: "), sg.Input(email_list.iloc[0].to_string(index=False), key="-email_1-")],
        [sg.Text("Email 2: "), sg.Input(email_list.iloc[1].to_string(index=False), key="-email_2-")],
        [sg.Text("Email 3: "), sg.Input(email_list.iloc[2].to_string(index=False), key="-email_3-")],
        [sg.Text("Email 4: "), sg.Input(email_list.iloc[3].to_string(index=False), key="-email_4-")],
        [sg.Button("Salvar"), sg.Button("Fechar")]
        ]
    
    window_ = sg.Window('Editar Lista Email', layout_lista,use_custom_titlebar=True, element_justification='c', titlebar_icon=favicon)
    while True:             # Event Loop
        event, values = window_.read()
        if event == sg.WIN_CLOSED or event == "Fechar":
            window_.close()
            break
        if event == "Salvar":
            try:       
                email_list.iloc[0] = values["-email_1-"]
                email_list.iloc[1] = values["-email_2-"]
                email_list.iloc[2] = values["-email_3-"]
                email_list.iloc[3] = values["-email_4-"]         
                email_list.to_csv("Data/Email.txt", index=False)
                sg.popup_quick_message("Salvo!")
            except: sg.popup_error("Ops! Algo deu errado")
    start_program()         
    
# Edit Product
def editar_produto():
    # consts
    df = pd.read_csv("Data/product_database.csv",sep="|")
    df = df.dropna(axis = 0, how = 'all')
    colunas = [sg.Text(i,background_color="black", size=15, font="Gadugi", justification="c", text_color="white") for i in df.columns.values]
    linhas = [[sg.Input(df.Produto.iloc[i], key=f"-produto_{i}-", size=20),sg.Input(df.volume.iloc[i], key=f"-volume_{i}-", size=20),sg.Input(df.Setor.iloc[i], key=f"-setor_{i}-", size=20)] for i in df.index]
    
    # layout
    layout_produtos = [
        [sg.Text("Lista de Produtos", font=('Helvetica', 13), justification="center")],
        colunas,
        [sg.Column(linhas, scrollable=True,  vertical_scroll_only=True,size=(450,200), justification="c")],
        [sg.Input(key="-prod-",size=(18)),sg.Input(key="-vol-",size=(18)),sg.Combo(["Injecao","Sopro"],key="-setor-",size=(18)),sg.Button("+")],
        [sg.Button("Salvar")]
        ]
    
    
    # Event Loop
    window_ = sg.Window('Lista de Produtos', layout_produtos,use_custom_titlebar=True, element_justification='c', titlebar_icon=favicon,size=(500,400))
    while True:              
        event, values = window_.read()
        if event == sg.WIN_CLOSED:
            break
        if event == "Salvar":
            try:
                for i in df.index:
                    df.Produto.iloc[i] = values["-produto_{}-".format(i)]     
                    df.volume.loc[i] = values["-volume_{}-".format(i)]     
                    df.Setor.iloc[i] = values["-setor_{}-".format(i)]
                df.dropna().to_csv("Data\product_database.csv", sep="|", index=False)
                sg.popup_quick_message("Salvo!")
            except: sg.popup_error("Ops! Algo deu errado")
            window_.close()
            editar_produto()
        if event == "+":
            if  values["-prod-"] != "" and values["-setor-"] != "":
                new = pd.DataFrame({"Produto":values["-prod-"],"volume":values["-vol-"],"Setor":values["-setor-"]}, index=[10])
                df = pd.concat([df,new], ignore_index=True).to_csv("Data\product_database.csv", sep="|", index=False)
                window_.close()
                editar_produto()
            else: 
                sg.PopupError("O nome e o setor são obrigatorios")
    start_program()        

# Pigmentos
def editar_pigmento():
    pigmento = pd.read_csv("Data/pigmento.csv", sep=";")
    pigmento = pigmento.dropna(axis = 0, how = 'all')
    colunas = [sg.Text(i,background_color="black", size=16, font="Gadugi", justification="c", text_color="white") for i in pigmento.columns.values]
    linhas = [[sg.Input(pigmento.cor.iloc[i], key=f"-cor_{i}-", size=20),sg.Input(pigmento.codigo.iloc[i], key=f"-codigo_{i}-", size=20)] for i in pigmento.index]
    
    layout_pig = [
        [sg.Text("Lista de Produtos", font=('Helvetica', 13), justification="center")],
        colunas,
        [sg.Column(linhas, scrollable=True, size=(300,200), vertical_scroll_only=True)],
        [sg.Input(key="-cor-",size=(20)),sg.Input(key="-cod-",size=(20)),sg.Button("+")],
        [sg.Button("Salvar")]
        ]
      
    window = sg.Window('Editar Lista Email', layout_pig,use_custom_titlebar=True, element_justification='c', titlebar_icon=favicon)
    while True:             # Event Loop
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == "Salvar":
            try:
                for i in pigmento.index:
                    pigmento.cor.iloc[i] = values["-cor_{}-".format(i)]     
                    pigmento.codigo.iloc[i] = values["-codigo_{}-".format(i)] 
                 
                pigmento.dropna().to_csv("Data\pigmento.csv", sep=";", index=False)
                sg.popup_quick_message("Salvo!")
            except: sg.popup_error("Ops! Algo deu errado")
            window.close()
            editar_pigmento()
        if event == "+":
            if  values["-cor-"] != "" and values["-cod-"] != "":
                new = pd.DataFrame({"cor":values["-cor-"], "codigo":values["-cod-"]}, index=[10])
                pigmento = pd.concat([pigmento,new], ignore_index=True).to_csv("Data\pigmento.csv", sep=";", index=False)
                window.close()
                editar_pigmento()  
            else: 
                sg.PopupError("Campos Vazios!")
    start_program()                 
    



start_program()
    
