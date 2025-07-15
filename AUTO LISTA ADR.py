from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile
from tkinter import messagebox
import os
import pygsheets
import platform

raiz = Tk()
raiz.geometry("800x600")
raiz.title('AUTO LISTA ADR')

escritas = 0
projeto = ''
div = ''
linha_widget = 0
saidas = []
importado = False
conectado = False
planilha = ''
pagina = ''
textos_atores = {}
coluna_TC = ''
coluna_personagem = ''
coluna_texto = ''
tecnica = ''
coluna_motivo = ''
coluna_obs = ''
coluna_tcout = ''
coluna_idioma = ''
coluna_rolo = ''
motivos = {'R': 'Ruidoso', 'L': 'Lapela falha', 'P': 'Pouca projeção, muito ruído', 'E': 'Estourado', 'S': 'Sem SD', 'F': 'Fala sobreposta', 'M': 'Consertar emenda da montagem', 'D': 'Mic distante'}


def escrever(celulas, valor, lin):
    # Função para escrever as informações no google sheets.
    # Ela recebe um tupla de células na seguinte ordem: TC IN,
    # Personagem, Texto, Motivo, Observação, Técnica, TC OUT, Idioma e Rolo.
    print(f'Aguarde, atualizando a linha {lin} da página.')
    global escritas
    escritas += 1
    
    # Lista que define quais colunas precisam ser formatadas. As 3 primeiras sempre são verdadeiras.
    atualizar = [True, True, True, False, False, False, False, False, False]
    pagina.update_value(celulas[0], valor['TC'])
    pagina.update_value(celulas[1], valor['Pers'])
    pagina.update_value(celulas[2], valor['Texto'])
    if valor['Motivo'] and coluna_motivo:
        pagina.update_value(celulas[3], valor['Motivo'])
        atualizar[3] = True
    if valor['OBS'] and coluna_obs:
        pagina.update_value(celulas[4], valor['OBS'])
        atualizar[4] = True
    if tecnica:
        pagina.update_value(celulas[5], 'Técnica')
        atualizar[5] = True
    if valor['TCOUT'] and coluna_tcout:
        pagina.update_value(celulas[6], valor['TCOUT'])
        atualizar[6] = True
    if coluna_idioma:
        pagina.update_value(celulas[7], valor['Idioma'])
        atualizar[7] = True
    if coluna_rolo:
        pagina.update_value(celulas[8], valor['Rolo'])
        atualizar[8] = True
    cont = 0
    for celula in celulas:
        if atualizar[cont]:
            # Se a célula atual foi modificada, formatar com WRAP, alinhamento centralizado e fonte Oswald
            pagina.cell(celula).wrap_strategy = 'WRAP'
            pagina.cell(celula).set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER)
            pagina.cell(celula).set_text_format('fontFamily', 'Oswald')
        cont += 1


def tc1_menor(tc1, tc2):
    # Essa é uma função recursiva que avalia se o primeiro timecode enviado é menor que o segundo.
    # Para isso, ela pega os dois primeiro caracteres do timecode (que tem formato
    # 00:00:00:00) e converte para inteiro, comparando ambos.
    tc1_num = int(tc1[:2])
    tc2_num = int(tc2[:2])
    if tc1_num < tc2_num:
        return True
    elif tc1_num == tc2_num:
        # Caso a primeira parte dos dois TCs sejam iguais, e função chama ela mesma para avaliar
        # a próxima parte do TC. Para isso ela envia para ela mesma o TC sem os 3 primeiros caracteres,
        # a não ser que ela já esteja nos últimos 2 caracteres. Nesse caso os TCs são exatamente iguais.
        if len(tc1) > 2:
            if tc1_menor(tc1[3:], tc2[3:]):
                return True
    return False


# Caracteres especiais e sua conversão para o unicode que o RTF consegue ler
especiais = {'À': '\\\'c0', 'Á': '\\\'c1', 'Â': '\\\'c2', 'Ã': '\\\'c3', 'Ç': '\\\'c7',
             'É': '\\\'c9', 'Ê': '\\\'ca', 'Í': '\\\'cd', 'Ó': '\\\'d3', 'Ô': '\\\'d4',
             'Õ': '\\\'d5', 'Ú': '\\\'da', 'à': '\\\'e0', 'á': '\\\'e1', 'â': '\\\'e2',
             'ã': '\\\'e3', 'ç': '\\\'e7', 'é': '\\\'e9', 'ê': '\\\'ea', 'í': '\\\'ed',
             'ó': '\\\'f3', 'ô': '\\\'f4', 'õ': '\\\'f5', 'ú': '\\\'fa'}


def converter_rtf(letra):
    global especiais
    # Caso especial em que o caractere é reticências (ord = 8230)
    if ord(letra) == 8230:
        return '\\\'85'
    if letra in especiais:
        return especiais[letra]
    return letra


def comparar_div(div1, div2):
    # Função que compara a divisão (episódio ou rolo), pegando o segundo elemento da lista gerada
    # a partir da string (exemplo: em EP 2, ela pega o 2), e convertendo para inteiro.
    div1_cortada = div1.split()
    div1_cortada[1] = int(div1_cortada[1])
    div2_cortada = div2.split()
    div2_cortada[1] = int(div2_cortada[1])
    if div1_cortada[1] < div2_cortada[1]:
        return 'menor'
    elif div1_cortada[1] > div2_cortada[1]:
        return 'maior'
    return 'igual'


def string_sem_lixo(string):
    # Retorna a string do marker sem o MARKERS RULER do fim
    controle = string.split()
    retorno = ''
    for palavra in controle:
        if palavra.upper() == 'MARKERS':
            return retorno
        else:
            retorno += f'{palavra} '
    return retorno


def valida_importar():
    # Habilita o botão de importar quando a opção 'Rolos Juntos' é escolhida
    # e nada ainda foi importado, ou quando um número foi inserido em txt_numero.
    # Desabilita esse botão se o botão rádio escolhido não for 'Rolos Juntos',
    # e não houver nenhum número em txt_numero.
    global btn_importar
    n = txt_numero.get()
    if div_var.get() == 'Rolos juntos':
        # Quando 'Rolos Juntos' é escolhido, txt_numero fica desabilitado
        txt_numero['state'] = 'disabled'
        if not importado:
            btn_importar['state'] = 'normal'
    else:
        if n:
            if int(n) > 0 and not importado:
                btn_importar['state'] = 'normal'
            else:
                btn_importar['state'] = 'disabled'
        else:
            btn_importar['state'] = 'disabled'
        txt_numero['state'] = 'normal'
    

def importar_txt():
    # Importa os markers válidos e armazena em saidas
    global div
    global saidas
    global importado
    div = div_var.get()
    if div != 'Rolos juntos':
        div += f' {txt_numero.get()}'
    path_txt = filedialog.askopenfilename(filetypes=[('Arquivos txt', '*.txt')])
    if path_txt:
        nome_txt = os.path.basename(path_txt)
        saidas = []
        invalidos = []
        arq = open(path_txt, 'r', encoding='utf-8')
        s = arq.readline()
        while s.rstrip() != 'M A R K E R S  L I S T I N G' and s != '':
            s = arq.readline()
        if s:
            arq.readline()
            s = arq.readline()
        while s != '':
            # A partir da linha correta, o programa avalia quais markers são válidos para a lista de ADR
            valores = s.split()
            marker = ''
            if len(valores) >= 5:
                marker = valores[4]
            if marker.upper() == 'ADR':
                tc_marker = valores[1]
                pers_marker = ''
                texto_marker = ''
                motivo_marker = ''
                obs_marker = ''
                idioma_marker = ''
                tcout_marker = ''
                check = 0
                for c in s:
                    if check == 0:
                        if c == 'a' or c == 'A':
                            check += 1
                    elif check == 1:
                        if c == 'd' or c == 'D':
                            check += 1
                        else:
                            check = 0
                    elif check == 2:
                        if c == 'r' or c == 'R':
                            check += 1
                        else:
                            check = 0
                    elif check == 3:
                        if c != '\"':
                            pers_marker += c
                        else:
                            pers_marker = pers_marker.strip().upper()
                            check += 1
                    elif check == 4:
                        if c != '\"':
                            texto_marker += c
                        else:
                            texto_marker = string_sem_lixo(texto_marker)
                            texto_marker = texto_marker.strip()  
                            check += 1
                    elif check == 5:
                        if c != '#':
                            motivo_marker += c
                        else:
                            check += 1
                    elif check == 6:
                        if c != '#':
                            obs_marker += c
                        else:
                            check += 1
                    elif check == 7:
                        if c != '#':
                            tcout_marker += c
                        else:
                            check += 1
                    else:
                        idioma_marker += c

                # Armazena as informações se o marker for válido
                if check > 4 and pers_marker and texto_marker:
                    motivo_marker = string_sem_lixo(motivo_marker)
                    motivo_marker = motivo_marker.strip()
                    obs_marker = string_sem_lixo(obs_marker)
                    obs_marker = obs_marker.strip()
                    tcout_marker = string_sem_lixo(tcout_marker)
                    tcout_marker = tcout_marker.strip()
                    idioma_marker = string_sem_lixo(idioma_marker)
                    idioma_marker = idioma_marker.strip()
                    if motivo_marker.upper() in motivos:
                        motivo_marker = motivos[motivo_marker.upper()]
                    if obs_marker.upper() == 'C':
                        obs_marker = 'Pode ser por celular'
                    if idioma_marker:
                        if idioma_marker == 'e' or idioma_marker == 'E':
                            idioma_marker = 'Español'
                    else:
                        idioma_marker = 'Português'
                    if div.split()[0] == 'Rolo':
                        rolo_marker = div.split()[1]
                    else:
                        rolo_marker = int(tc_marker[:2])
                    saidas.append({'TC': tc_marker, 'Pers': pers_marker, 'Texto': texto_marker,
                                   'Motivo': motivo_marker, 'OBS': obs_marker,
                                   'TCOUT': tcout_marker, 'Idioma': idioma_marker, 'Rolo': rolo_marker})
                else:
                    invalidos.append(s)
            else:
                invalidos.append(s)
            s = arq.readline()
        arq.close()
        if len(saidas) > 0:
            # Impede que a divisão e número do projeto sejam alterados 
            importado = True
            btn_apagar['state'] = 'normal'
            lbl_importado['text'] = f'{len(saidas)} marker(s) importado(s) de {nome_txt}'
            lbl_importado['fg'] = 'green'
            txt_numero['state'] = 'disabled'
            rad_ep['state'] = 'disabled'
            rad_rolo['state'] = 'disabled'
            rad_rolos['state'] = 'disabled'
            btn_executar['state'] = 'normal'
        else:
            lbl_importado['text'] = f'Nenhum marker importado (não foram encontrados markers válidos em {nome_txt})'
            lbl_importado['fg'] = 'red'
        if len(invalidos) > 0:
            print('Lista de markers inválidos:\n')
            for invalido in invalidos:
                print(invalido)


def apagar_markers():
    # Limpa os dados armazenados em saidas e permite que a divisão e número
    # do projeto sejam alterados
    global saidas
    global importado
    saidas = []
    importado = False
    valida_importar()
    btn_apagar['state'] = 'disabled'
    lbl_importado['text'] = 'Nenhum marker importado (markers apagados pelo usuário)'
    lbl_importado['fg'] = 'black'
    if div_var.get() != 'Rolos juntos':
        txt_numero['state'] = 'normal'
    rad_ep['state'] = 'normal'
    rad_rolo['state'] = 'normal'
    rad_rolos['state'] = 'normal'
    if not conectado:
        btn_executar['state'] = 'disabled'
        

def numero_valido(string, nova):
    # Regula a entrada de números (tamanho máximo de 3, só algorismos)
    if len(nova) > 3:
        return False
    return string.isdecimal()


def coluna_valida(string, nova):
    # Regula a entrada de colunas (tamanho máximo de 1, só letras)
    if len(nova) > 1:
        return False
    return string.isalpha()


def conectar():
    # Conecta-se com uma planilha Google
    pl = txt_link.get()
    pag = txt_pagina.get()
    global conectado
    if pl and pag:
        lbl_status['fg'] = 'black'
        global planilha
        try:
            gc = pygsheets.authorize()
            try:
                planilha = gc.open_by_url(pl)
                global pagina
                i = 0
                for p in planilha:
                    if pag == p.title:
                        pagina = planilha[i]
                        break
                    i += 1
                if i == len(planilha.worksheets()):
                    lbl_status['text'] = 'Não foi possível encontrar a página da planilha. Verifique se o nome está correto'
                    lbl_status['fg'] = 'red'
                    conectado = False
                else:
                    lbl_status['text'] = f'Conectado à planilha {planilha.title}, página {pagina.title}'
                    lbl_status['fg'] = 'green'
                    conectado = True
                    btn_executar['state'] = 'normal'
                    btn_desconectar['state'] = 'normal'
            except:
                lbl_status['text'] = 'Não foi possível encontrar a planilha. Verifique se o link está correto'
                lbl_status['fg'] = 'red'
                conectado = False
        except:
            lbl_status['text'] = 'Não foi possível conseguir autorização. Checar se o token está expirado.'
            lbl_status['fg'] = 'red'
            conectado = False
    else:
        lbl_status['text'] = 'É necessário informar o link da planilha e o nome da página'
        lbl_status['fg'] = 'red'


def desconectar():
    # Desconecta-se da planilha Google
    lbl_status['text'] = 'Desconectado.'
    lbl_status['fg'] = 'black'
    conectado = False
    if not importado:
        btn_executar['state'] = 'disabled'
    btn_desconectar['state'] = 'disabled'


def exportar_textos():
    # Exporta textos em formato rtf com as informações contidas em texto_atores
    exportados = ''
    messagebox.showinfo('Informação', 'O progama vai pedir para você escolher uma pasta para salvar os textos gerados.')
    local_saida = filedialog.askdirectory()
    if local_saida:
        for personagem in textos_atores:
            # Criação dos arquivos de texto dos atores na pasta indicada pelo usuário
            nome_arquivo = f'{projeto.upper()}_ADR_{personagem}.rtf'
            path_arquivo = os.path.join(local_saida, nome_arquivo)
            saida_anterior = ''
            saida_posterior = ''
            igual = False
            if div != 'Rolos juntos':
                # Checar se já existe arquivo com esse nome na mesma pasta e, se sim,
                # ler e guardar o que tiver nele em episódios ou rolos anteriores ao atual,
                # e também nos posteriores.
                if os.path.exists(path_arquivo):
                    arq = open(path_arquivo, 'r')
                    s = arq.readline()
                    div_saida = div.split()
                    div_lida = ''
                    while div_lida != div_saida[0] and s != '':
                        s = arq.readline()
                        if len(s) >= len(div_saida[0]):
                            div_lida = s[:len(div_saida[0])]
                    if s:
                        terminou = False
                        while not terminou:
                            div_lida = ''
                            if comparar_div(s, div) == 'menor':
                                while div_lida != div_saida[0] and s != '':
                                    if s != '}':
                                        saida_anterior += s
                                    s = arq.readline()
                                    if len(s) >= len(div_saida[0]):
                                        div_lida = s[:len(div_saida[0])]
                            elif comparar_div(s, div) == 'maior':
                                while div_lida != div_saida[0] and s != '':
                                    saida_posterior += s
                                    s = arq.readline()
                                    if len(s) >= len(div_saida[0]):
                                        div_lida = s[:len(div_saida[0])]
                            else:
                                igual = True
                                while div_lida != div_saida[0] and s != '':
                                    s = arq.readline()
                                    if len(s) >= len(div_saida[0]):
                                        div_lida = s[:len(div_saida[0])]
                            if not s:
                                terminou = True

                    arq.close()
            arq = open(path_arquivo, 'w')
            # Escrever conforme formatação do arquivo RTF
            arq.write('{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Arial;}}\n'
                      '{\\pard\\qc\\f0\\fs36\\b\\sl270\\slmult1\n')
            texto_final = ''
            for c in personagem:
                texto_final += converter_rtf(c)
            arq.write(f'ADR {texto_final}\n\\line\\line\n')
            if saida_anterior:
                arq.write(saida_anterior)
            if saida_anterior and not saida_posterior and not igual:
                arq.write('{\\pard\\qc\\f0\\fs36\\b\\sl270\\slmult1\n')
            if div != 'Rolos juntos':
                arq.write(f'{div}\n\\line\\line\n')
            else:
                menor_tc = ''
                for timecode in textos_atores[personagem]:
                    if menor_tc:
                        if tc1_menor(timecode, menor_tc):
                            menor_tc = timecode
                    else:
                        menor_tc = timecode
                rolo = int(menor_tc[:2])
                arq.write(f'Rolo {rolo}\n\\line\\line\n')
            arq.write('\\par}\n{\\pard\\fs36\\sl270\\slmult1\n')
            for timecode in textos_atores[personagem]:
                if div == 'Rolos juntos':
                    if int(timecode[:2]) != rolo:
                        rolo = int(timecode[:2])
                        arq.write('\\par}\n')
                        arq.write('{\\pard\\qc\\f0\\fs36\\b\\sl270\\slmult1\n')
                        arq.write(f'Rolo {rolo}\n\\line\\line\n')
                        arq.write('\\par}\n{\\pard\\fs36\\sl270\\slmult1\n')
                arq.write(timecode)
                arq.write(f'\n\\line\\line\n{textos_atores[personagem][timecode]}\n\\line\\line\\line\n')
            arq.write('\\par}\n')
            if saida_posterior:
                arq.write('{\\pard\\qc\\f0\\fs36\\b\\sl270\\slmult1\n')
                arq.write(saida_posterior)
            else:
                arq.write('}')
            arq.close()
            exportados += f'{nome_arquivo}\n'
        if len(textos_atores) == 0:
            lbl_gerados['text'] = 'Não havia informações para carregar e nenhum texto foi gerado.'
            lbl_gerados['fg'] = 'red'
        else:
            lbl_gerados['text'] = f'O(s) seguinte(s) arquivo(s) de texto foi(foram) gerado(s) com sucesso:\n{exportados}'
            lbl_gerados['fg'] = 'green'


def executar():
    # Preenche a planilha Google, se houver conexão, com os dados armazenados em saidas,
    # se houve importação de markers, e no fim chama a função de exportar textos com
    # os dados armazenados ou dos markers, ou da planilha, ou de ambos, conforme opção do
    # usuário
    global textos_atores
    global projeto
    global div
    projeto = txt_projeto.get()
    div = div_var.get()
    if div != 'Rolos juntos':
        div += f' {txt_numero.get()}'
    if projeto:
        if txt_numero.get() or div == 'Rolos juntos':
            if conectado:
                global coluna_TC
                global coluna_personagem
                global coluna_texto
                global tecnica
                global coluna_motivo
                global coluna_obs
                global coluna_tcout
                global coluna_idioma
                global coluna_rolo
                coluna_TC = txt_tc.get()
                coluna_personagem = txt_pers.get()
                coluna_texto = txt_texto.get()
                linha = txt_linha.get()
                if coluna_TC and coluna_personagem and coluna_texto and linha:
                    coluna_TC += f'{linha}'
                    coluna_personagem += f'{linha}'
                    coluna_texto += f'{linha}'
                    tecnica = txt_tecnica.get()
                    coluna_motivo = txt_motivo.get()
                    coluna_obs = txt_obs.get()
                    coluna_tcout = txt_tcout.get()
                    coluna_idioma = txt_idioma.get()
                    coluna_rolo = txt_rolo.get()
                    linha_final = txt_linha_final.get()
                    linha = int(linha)
                    if not linha_final:
                        linha_final = linha - 1
                    else:
                        linha_final = int(linha_final)
                    if tecnica:
                        tecnica += f'{linha}'
                    if coluna_motivo:
                        coluna_motivo += f'{linha}'
                    if coluna_obs:
                        coluna_obs += f'{linha}'
                    if coluna_tcout:
                        coluna_tcout += f'{linha}'
                    if coluna_idioma:
                        coluna_idioma += f'{linha}'
                    if coluna_rolo:
                        coluna_rolo += f'{linha}'
                    TC_atual = coluna_TC
                    pers_atual = coluna_personagem
                    texto_atual = coluna_texto
                    motivo_atual = coluna_motivo
                    obs_atual = coluna_obs
                    tecnica_atual = tecnica
                    tcout_atual = coluna_tcout
                    idioma_atual = coluna_idioma
                    rolo_atual = coluna_rolo
                    for saida in saidas:
                        valido = False
                        while not valido:
                            print(f'Aguarde, checando a linha {linha} da página.')
                            TC_pagina = pagina.get_value(TC_atual).strip()
                            pers_pagina = pagina.get_value(pers_atual).upper().strip()
                            texto_pagina = pagina.get_value(texto_atual).strip()
                            # As strings a seguir são as que serão escritas nos textos dos atores. Só serão usadas
                            # TC, personagem e texto.
                            pers_para_texto = ''
                            tc_para_texto = ''
                            texto_para_texto = ''
                            completo = False
                            if TC_pagina == '':
                                if pers_pagina == '' and texto_pagina == '':
                                    escrever((TC_atual, pers_atual, texto_atual,
                                              motivo_atual, obs_atual, tecnica_atual, tcout_atual,
                                              idioma_atual, rolo_atual), saida, linha)
                                    pers_para_texto = saida['Pers']
                                    tc_para_texto = saida['TC']
                                    texto_para_texto = saida['Texto']
                                    valido = True
                                    completo = True
                            else:
                                if tc1_menor(saida['TC'], TC_pagina):
                                    pagina.insert_rows(row=linha-1, number=1)
                                    escrever((TC_atual, pers_atual, texto_atual,
                                              motivo_atual, obs_atual, tecnica_atual, tcout_atual,
                                              idioma_atual, rolo_atual), saida, linha)
                                    pers_para_texto = saida['Pers']
                                    tc_para_texto = saida['TC']
                                    texto_para_texto = saida['Texto']
                                    valido = True
                                    completo = True
                                elif saida['TC'] == TC_pagina:
                                    if saida['Texto'] != texto_pagina or saida['Pers'] != pers_pagina:
                                        if saida['Pers'] == pers_pagina:
                                            novo_tc = TC_pagina[:-2]
                                            fim = int(TC_pagina[-2:])
                                            fim += 1
                                            if fim > 9:
                                                fim = str(fim)
                                            else:
                                                fim = f'0{fim}'
                                            novo_tc += fim
                                            pagina.update_value(TC_atual, novo_tc)
                                        pagina.insert_rows(row=linha-1, number=1)
                                        escrever((TC_atual, pers_atual, texto_atual,
                                                  motivo_atual, obs_atual, tecnica_atual,
                                                  tcout_atual, idioma_atual, rolo_atual), saida, linha)
                                    pers_para_texto = saida['Pers']
                                    tc_para_texto = saida['TC']
                                    texto_para_texto = saida['Texto']
                                    valido = True
                                    completo = True
                                else:
                                    if pers_pagina != '' and texto_pagina != '':
                                        pers_para_texto = pers_pagina
                                        tc_para_texto = TC_pagina
                                        texto_para_texto = texto_pagina
                                        completo = True
                            if completo:
                                texto_final = ''
                                for c in texto_para_texto:
                                    texto_final += converter_rtf(c)
                                texto_para_texto = texto_final
                                if pers_para_texto not in textos_atores:
                                    textos_atores[pers_para_texto] = {}
                                textos_atores[pers_para_texto][tc_para_texto] = texto_para_texto
                            linha += 1
                            TC_atual = f'{TC_atual[:1]}{linha}'
                            pers_atual = f'{pers_atual[:1]}{linha}'
                            texto_atual = f'{texto_atual[:1]}{linha}'
                            if coluna_motivo:
                                motivo_atual = f'{motivo_atual[:1]}{linha}'
                            if coluna_obs:
                                obs_atual = f'{obs_atual[:1]}{linha}'
                            if tecnica:
                                tecnica_atual = f'{tecnica_atual[:1]}{linha}'
                            if coluna_tcout:
                                tcout_atual = f'{tcout_atual[:1]}{linha}'
                            if coluna_idioma:
                                idioma_atual = f'{idioma_atual[:1]}{linha}'
                            if coluna_rolo:
                                rolo_atual = f'{rolo_atual[:1]}{linha}'

                    linha_final += escritas
                    
                    while linha <= linha_final:
                        # Armazenar as informações das linhas restantes, se houver, para o texto final.
                        print(f'Aguarde, checando a linha {linha} da página.')
                        TC_pagina = pagina.get_value(TC_atual).strip()
                        pers_pagina = pagina.get_value(pers_atual).upper().strip()
                        texto_pagina = pagina.get_value(texto_atual).strip()
                        pers_para_texto = ''
                        tc_para_texto = ''
                        texto_para_texto = ''
                        if TC_pagina != '' and pers_pagina != '' and texto_pagina != '':
                            pers_para_texto = pers_pagina
                            tc_para_texto = TC_pagina
                            texto_para_texto = texto_pagina
                            texto_final = ''
                            for c in texto_para_texto:
                                texto_final += converter_rtf(c)
                            texto_para_texto = texto_final
                            if pers_para_texto not in textos_atores:
                                textos_atores[pers_para_texto] = {}
                            textos_atores[pers_para_texto][tc_para_texto] = texto_para_texto
                        linha += 1
                        TC_atual = f'{TC_atual[:1]}{linha}'
                        pers_atual = f'{pers_atual[:1]}{linha}'
                        texto_atual = f'{texto_atual[:1]}{linha}'
                        if coluna_motivo:
                            motivo_atual = f'{motivo_atual[:1]}{linha}'
                        if coluna_obs:
                            obs_atual = f'{obs_atual[:1]}{linha}'
                        if tecnica:
                            tecnica_atual = f'{tecnica_atual[:1]}{linha}'
                        if coluna_tcout:
                            tcout_atual = f'{tcout_atual[:1]}{linha}'
                        if coluna_idioma:
                            idioma_atual = f'{idioma_atual[:1]}{linha}'
                        if coluna_rolo:
                            rolo_atual = f'{rolo_atual[:1]}{linha}'
                            
                    exportar_textos()
                else:
                    messagebox.showwarning('Atenção', 'Você não preencheu todas as informações obrigatórias da planilha')
            else:
                for saida in saidas:
                    pers_para_texto = saida['Pers']
                    tc_para_texto = saida['TC']
                    texto_para_texto = saida['Texto']
                    texto_final = ''
                    for c in texto_para_texto:
                        texto_final += converter_rtf(c)
                    texto_para_texto = texto_final
                    if pers_para_texto not in textos_atores:
                        textos_atores[pers_para_texto] = {}
                    textos_atores[pers_para_texto][tc_para_texto] = texto_para_texto
                exportar_textos()
        else:
            messagebox.showwarning('Atenção', f'Você precisa informar o número do {div_var.get()}')
    else:
        messagebox.showwarning('Atenção', 'Você precisa informar a sigla do projeto')


def abrir():
    # Abre um projeto salvo e usa suas informações nos widgets
    path_txt = filedialog.askopenfilename(filetypes=[('Arquivos txt', '*.txt')])
    if path_txt:
        arq = open(path_txt, 'r', encoding='utf-8')
        s = arq.readline()
        if s == 'Arquivo de projeto salvo do programa Auto Lista ADR\n':
            s = arq.readline()
            i = 1
            while s:
                entrada = s.rstrip()
                if i == 1:
                    txt_projeto.delete(0,END)
                    txt_projeto.insert(0,entrada)
                elif i == 2:
                    div_var.set(entrada)
                elif i == 3:
                    txt_link.delete(0,END)
                    txt_link.insert(0,entrada)
                elif i == 4:
                    txt_pagina.delete(0,END)
                    txt_pagina.insert(0,entrada)
                elif i == 5:
                    txt_tc.delete(0,END)
                    txt_tc.insert(0,entrada)
                elif i == 6:
                    txt_pers.delete(0,END)
                    txt_pers.insert(0,entrada)
                elif i == 7:
                    txt_texto.delete(0,END)
                    txt_texto.insert(0,entrada)
                elif i == 8:
                    txt_motivo.delete(0,END)
                    txt_motivo.insert(0,entrada)
                elif i == 9:
                    txt_obs.delete(0,END)
                    txt_obs.insert(0,entrada)
                elif i == 10:
                    txt_tecnica.delete(0,END)
                    txt_tecnica.insert(0,entrada)
                elif i == 11:
                    txt_tcout.delete(0,END)
                    txt_tcout.insert(0,entrada)
                elif i == 12:
                    txt_idioma.delete(0,END)
                    txt_idioma.insert(0,entrada)
                elif i == 13:
                    txt_rolo.delete(0,END)
                    txt_rolo.insert(0,entrada)
                elif i == 14:
                    txt_linha.delete(0,END)
                    txt_linha.insert(0,entrada)
                i += 1
                s = arq.readline()
        else:
            os.path.basename(path_txt)
            messagebox.showwarning('Atenção', f'{os.path.basename(path_txt)} não é um arquivo válido de projeto Auto Lista ADR')
        arq.close()


def salvar(proj):
    # Salva um projeto com o conteúdodos widgets
    nome_saida = proj
    if not nome_saida:
        nome_saida = 'semtitulo'
    arq = asksaveasfile(initialfile = f'{nome_saida}_ALA.txt', filetypes=[("Text Documents","*.txt")])
    if arq:
        arq.write(f'Arquivo de projeto salvo do programa Auto Lista ADR\n'
                  f'{proj}\n{div_var.get()}\n{txt_link.get()}\n{txt_pagina.get()}\n{txt_tc.get()}'
                  f'\n{txt_pers.get()}\n{txt_texto.get()}\n{txt_motivo.get()}\n{txt_obs.get()}\n{txt_tecnica.get()}'
                  f'\n{txt_tcout.get()}\n{txt_idioma.get()}\n{txt_rolo.get()}\n{txt_linha.get()}')
        arq.close()
        

padx_1 = 5
pady_0 = 5
pady_1 = (pady_0,45)

# Inicialização e posicionamento de cada widget do programa

# Canvas e Scrollbar
canvas = Canvas(raiz, borderwidth=0, highlightthickness=0)
scrollbar = Scrollbar(raiz, orient=VERTICAL, command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollbar.pack(side=RIGHT, fill=Y)
canvas.pack(side=LEFT, fill=BOTH, expand=True)

# Frame dentro do canvas para conter os widgets
frame = Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor="nw")

# Atualiza a scrollregion quando o frame for redimensionado
def on_frame_config(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

frame.bind("<Configure>", on_frame_config)

# Roda do mouse
def _on_mousewheel(event):
    if platform.system() == 'Windows':
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    elif platform.system() == 'Darwin':  # macOS
        canvas.yview_scroll(int(-1*(event.delta)), "units")
    else:  # Linux
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

# Windows e macOS
canvas.bind_all("<MouseWheel>", _on_mousewheel)

# Linux (usa botões 4 e 5 para rolagem)
canvas.bind_all("<Button-4>", _on_mousewheel)
canvas.bind_all("<Button-5>", _on_mousewheel)

txt_projeto = Entry(frame, width=8)
btn_abrir = Button(frame, text='Abrir projeto', command=abrir)
btn_abrir.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_1)
btn_salvar = Button(frame, text='Salvar projeto', command=lambda:salvar(txt_projeto.get()))
btn_salvar.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_1, columnspan=3)

linha_widget += 1

lbl_projeto = Label(frame, text='Sigla do projeto:')
lbl_projeto.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
txt_projeto.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

div_var = StringVar()
lbl_div = Label(frame, text='Escolha o tipo de divisão do projeto:')
lbl_div.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
rad_ep = Radiobutton(frame, text='Episódio', variable=div_var, value='EP', command=lambda:valida_importar())
linha_widget += 1
rad_ep.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
rad_rolo = Radiobutton(frame, text='Rolo', variable=div_var, value='Rolo', command=lambda:valida_importar())
rad_rolo.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0)
rad_rolos = Radiobutton(frame, text='Rolos juntos', variable=div_var, value='Rolos juntos', command=lambda:valida_importar())
rad_rolos.grid(row=linha_widget, column=2, sticky='W', padx=padx_1, pady=pady_0)
div_var.set('EP')

linha_widget += 1

lbl_numero = Label(frame, text='Número do EP ou do Rolo:')
lbl_numero.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_1)
txt_numero = Entry(frame, width=2, state='normal', validate='key', validatecommand=(frame.register(numero_valido), '%S', '%P'))
txt_numero.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_1)
txt_numero.bind('<KeyRelease>', lambda e: valida_importar())

linha_widget += 1

btn_importar = Button(frame, text='Importar Markers', state='disabled', command=importar_txt)
btn_importar.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
btn_apagar = Button(frame, text='Apagar markers importados', state='disabled', command=apagar_markers)
btn_apagar.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0, columnspan=2)

linha_widget += 1

lbl_importado = Label(frame, text='Nenhum marker importado (o usuário não importou)')
lbl_importado.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, columnspan=3, pady=pady_1)

linha_widget += 1

lbl_link = Label(frame, text='Link da planilha:')
lbl_link.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
txt_link = Entry(frame, width=50)
txt_link.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0, columnspan=4)

linha_widget += 1

lbl_pagina = Label(frame, text='Nome da página da planilha:')
lbl_pagina.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
txt_pagina = Entry(frame, width=15)
txt_pagina.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0, columnspan=3)

linha_widget += 1

btn_conectar = Button(frame, text='Conectar', command=conectar)
btn_conectar.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
btn_desconectar = Button(frame, text='Desconectar', state='disabled', command=desconectar)
btn_desconectar.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

lbl_status = Label(frame, text='Não houve conexão com nenhuma planilha')
lbl_status.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_1, columnspan=3)

linha_widget += 1

lbl_colunas = Label(frame, text='Informe a letra das colunas dos dados abaixo como estão na planilha '
                    '(as colunas com * são obrigatórias):')
lbl_colunas.grid(row=linha_widget, column=0, columnspan=5, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1


lbl_tc = Label(frame, text='TC IN*')
lbl_tc.grid(row=linha_widget, column=0, sticky='E', padx=padx_1)
txt_tc = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_tc.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0)
lbl_pers = Label(frame, text='Personagem*')
lbl_pers.grid(row=linha_widget, column=2, sticky='E', padx=padx_1, pady=pady_0)
txt_pers = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_pers.grid(row=linha_widget, column=3, sticky='W', padx=padx_1, pady=pady_0)
lbl_texto = Label(frame, text='Texto*')
lbl_texto.grid(row=linha_widget, column=4, sticky='E', padx=padx_1, pady=pady_0)
txt_texto = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_texto.grid(row=linha_widget, column=5, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

lbl_motivo = Label(frame, text='Motivo')
lbl_motivo.grid(row=linha_widget, column=0, sticky='E', padx=padx_1, pady=pady_0)
txt_motivo = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_motivo.grid(row=linha_widget, column=1,sticky='W', padx=padx_1, pady=pady_0)
lbl_obs = Label(frame, text='OBS')
lbl_obs.grid(row=linha_widget, column=2, sticky='E', padx=padx_1, pady=pady_0)
txt_obs = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_obs.grid(row=linha_widget, column=3, sticky='W', padx=padx_1)
lbl_tecnica = Label(frame, text='Artística/Técnica')
lbl_tecnica.grid(row=linha_widget, column=4, sticky='E', padx=padx_1, pady=pady_0)
txt_tecnica = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_tecnica.grid(row=linha_widget, column=5, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

lbl_tcout = Label(frame, text='TC OUT')
lbl_tcout.grid(row=linha_widget, column=0, sticky='E', padx=padx_1, pady=pady_1)
txt_tcout = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_tcout.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_1)
lbl_idioma = Label(frame, text='Idioma')
lbl_idioma.grid(row=linha_widget, column=2, sticky='E', padx=padx_1, pady=pady_1)
txt_idioma = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_idioma.grid(row=linha_widget, column=3, sticky='W', padx=padx_1, pady=pady_1)
lbl_rolo = Label(frame, text='Rolo')
lbl_rolo.grid(row=linha_widget, column=4, sticky='E', padx=padx_1, pady=pady_1)
txt_rolo = Entry(frame, width=1, validate='key', validatecommand=(frame.register(coluna_valida), '%S', '%P'))
txt_rolo.grid(row=linha_widget, column=5, sticky='W', padx=padx_1, pady=pady_1)

linha_widget += 1

lbl_linha = Label(frame, text='Linha de início das entradas na planilha*')
lbl_linha.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)
txt_linha = Entry(frame, width=2, validate='key', validatecommand=(frame.register(numero_valido), '%S', '%P'))
txt_linha.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

lbl_linha_final = Label(frame, text='Linha em que terminam as entradas:')
lbl_linha_final.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, columnspan=2, pady=pady_0)
txt_linha_final = Entry(frame, width=2, validate='key', validatecommand=(frame.register(numero_valido), '%S', '%P'))
txt_linha_final.grid(row=linha_widget, column=1, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

lbl_linha_final2 = Label(frame, text='(se não houver nenhuma entrada na planilha, não preencha)')
lbl_linha_final2.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=(0, 45), columnspan=2)

linha_widget += 1

btn_executar = Button(frame, text='Executar', state='disabled', command=executar)
btn_executar.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_0)

linha_widget += 1

lbl_gerados = Label(frame, text='Nenhum texto gerado')
lbl_gerados.grid(row=linha_widget, column=0, sticky='W', padx=padx_1, pady=pady_1)

raiz.mainloop()
