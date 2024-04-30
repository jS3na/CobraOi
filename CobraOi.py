import traceback
import pyautogui as pg
from time import sleep, time
from random import randint
from pynput.keyboard import Controller, Key
from pyperclip import copy
import tkinter as tk
from tkinter import filedialog
from pandas import read_excel, notna
import keyboard

def abrir_arquivo(): #para permitir que o usuario possa abrir a tabela do excel que quiser
    root = tk.Tk()
    root.withdraw()  # oculta a janela principal do tkinter
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls;*.csv")])
    return file_path

def devendoEconta():
    
    def Extrato():
        global count_extrato
        try:
            extrato = pg.locateOnScreen(r'CobraOi\imagens\extrato.png', region=(610, 150, 160, 700), confidence=0.88)
            if extrato:
                extrato = list(pg.locateAllOnScreen(r'CobraOi\imagens\extrato.png', region=(610, 150, 160, 700), confidence=0.88))
                count_extrato = len(extrato)
                print("extratos: {}".format(count_extrato))
        except pg.ImageNotFoundException:
            pass

    def Boleto():
        global count_boleto
        try:
            boleto = pg.locateOnScreen(r'CobraOi\imagens\boleto.png', region=(610, 150, 160, 700), confidence=0.88)
            if boleto:
                boleto = list(pg.locateAllOnScreen(r'CobraOi\imagens\boleto.png', region=(610, 150, 160, 700), confidence=0.88))
                count_boleto = len(boleto)
                print("boletos: {}".format(count_boleto))
        except pg.ImageNotFoundException:
            pass

    Extrato()
    Boleto()

    try:
        if count_boleto == count_extrato or count_extrato > count_boleto:
            return False
        elif count_boleto > count_extrato:
            print('devendo')
            return True
    except NameError:
        print("deu erro")

def verifOifibra(): #verificação se o serviço atual é apenas OiFibra
    
    global simOifibras
    
    try:
        busca = pg.locateOnScreen(r'CobraOi\imagens\oifibra.png', region=(470, 500, 500, 100), confidence = 0.7)
        
        if busca: 
            simOifibras = True
            return True
        
    except pg.ImageNotFoundException:
            simOifibras = False
            return False

def fixo(): #verificação se o serviço atual é apenas Fixo, o que não importa para a gente
    try:
        busca = pg.locateOnScreen(r'CobraOi\imagens\cancelado.png', region=(600, 600, 300, 100), confidence = 0.7)     
        if busca: return True
        
    except pg.ImageNotFoundException:    
            return False 

def telaInsereID(numeroID, opcoesIMG, opcoes, op): #responsável por toda a tela inicial
    
    if len(numeroID) <= 11: #se o número for cpf, ele preenche os respectivos campos
        
        sleep(1)
        txt = pg.locateOnScreen(r'CobraOi\imagens\nome_vendedor.png', region=(300, 200, 500, 300), confidence=0.9)
        x, y = pg.center(txt)
        pg.click(x, y+20, duration=0.5)
        sleep(0.3)
        pg.hotkey('ctrl', 'a')
        sleep(0.3)
        copy("TR747319")
        pg.hotkey('ctrl', 'v')
        pg.press("tab")
        
        sleep(1)
        
        txt = pg.locateOnScreen(r'CobraOi\imagens\cpf_cliente.png', region=(300, 250, 500, 300), confidence=0.9)
        x, y = pg.center(txt)
        pg.click(x, y+20, duration=0.8)
        
    else: #se o número for cnpj, ele preenche os respectivos campos
        
        sleep(1)
        txt = pg.locateOnScreen(r'CobraOi\imagens\cnpj_cliente.png', region=(300, 250, 500, 300), confidence=0.9)
        x, y = pg.center(txt)
        pg.click(x, y+20, duration=0.8)
        
    copy(numeroID)
    sleep(0.5)
    pg.hotkey('ctrl', 'v')
    
    pg.click(922,469, duration=1)
    sleep(6)
    
    try: #tenta verificar se a imagem "novocliente" existe (aparece quando o cliente foi cancelado)
        
        verificarCancel = pg.locateOnScreen(r'CobraOi\imagens\novocliente.png', region=(830, 450, 240, 100), confidence = 0.7)
        
        if verificarCancel:
            return True
        
    except pg.ImageNotFoundException: #escolhe as opções caso nao seja cliente cancelado
        
        verifOifibra()
        
        for ft in opcoesIMG: #pesquisa pelas 2 variações de opção
            try:
                
                if len(opcoes) != 0: #se em uma iteração passada já tiver agregado opções, não entra na verificação novamente
                    pass
                else:
                    
                    casos = pg.locateOnScreen(ft, region=(480, 430, 490, 240), confidence=0.7) #procura uma opção
                    
                    if casos:
                        casos = pg.locateAllOnScreen(ft, region=(480, 430, 490, 240), confidence=0.7) #procura todos os casos dessa opção do loop
                        for caso in casos:
                            opcoes.append(pg.center(caso)) #adiciona as aparições das opções (se for mais de 1)

                    print(opcoes)
                    
            except pg.ImageNotFoundException: continue
            
            finally: pass
            
            pg.moveTo(opcoes[op][0], opcoes[op][1], duration=0.5) #clica na respectiva opção da vez do loop
            sleep(1)
            pg.click()
            #print(oifibra())
            
            sleep(1)
            
            btt = pg.locateOnScreen(r'CobraOi\imagens\avancar_inicial.png', region=(700, 200, 300, 650), confidence=0.7)
            x, y = pg.center(btt)
            pg.click(x, y)
            
            break
        
    return False

def telaIniciaAtendimento(): #responsável pela segunda tela
    
    btt = pg.locateOnScreen(r'CobraOi\imagens\servicos_oi.png', region=(885, 100, 280, 790), confidence = 0.7) #BOTAO SERVIÇOS OI
    x, y = pg.center(btt)
    pg.click(x, y)
    
    sleep(0.5)
    pg.scroll(-500)
    sleep(0.8)
    
    btt = pg.locateOnScreen(r'CobraOi\imagens\iniciar_atendimento.png', region=(915, 100, 240, 790), confidence = 0.7) #BOTAO INICIAR ATENDIMENTO
    x, y = pg.center(btt)
    pg.click(x, y, duration=0.4)

def telaSegundaEcontas(): #terceira tela
    
    sleep(4)
    
    if not simOifibras:
        btt = pg.locateOnScreen(r'CobraOi\imagens\segunda_via.png', region=(500, 500, 500, 300), confidence = 0.7) #BOTAO SEGUNDA VIA
        x, y = pg.center(btt)
        pg.click(x, y, duration=0.5)
    
    else:
        btt = pg.locateOnScreen(r'CobraOi\imagens\econtas.png', region=(500, 500, 500, 300), confidence = 0.7) #BOTAO E-CONTAS
        x, y = pg.center(btt)
        pg.click(x, y, duration=0.5)

def selectFaturasMandar(teclado): #responsável por enviar as faturas para o cliente pela pasta de downloads
    
    btt = list(pg.locateAllOnScreen(r'CobraOi\imagens\downloads.png', confidence=0.7)) #BOTÃO DOWNLOADS
    x, y= pg.center(btt[0])
    pg.click(x, y, duration=0.5)
    sleep(1)
    pg.press("tab")
    sleep(0.3)
    pg.press('down')
    sleep(0.3)
    pg.press("up")
    sleep(0.5)
    
    if not simOifibras:
        if count_fatura > 1:
            
            with teclado.pressed(Key.shift):
                sleep(1)
                
                for _ in range(count_fatura-1):
                    sleep(1)
                    teclado.press(Key.down)
        
    pg.press("enter")
    sleep(5)
    pg.press("enter")
    sleep(1)
    pg.hotkey("ctrl", "1")

def msgWhatsapp(num, frases, teclado, nome): #responsável pela msg do whatsapp
    
    global achouNumero
    
    print('- Enviando as fatura(s) ao número {}1!'.format(num))
    sleep(2)
    pg.press('esc', presses=2, interval=0.8)
    sleep(1)
    
    btt = pg.locateOnScreen(r'CobraOi\imagens\novaconversa.png', region=(20, 80, 490, 240), confidence=0.7) #BOTÃO DE NOVA CONVERSA
    x, y = pg.center(btt)
    pg.click(x, y, duration=0.5)
    sleep(1)
    
    teclado.type(num)
    sleep(3)
    pg.press("enter")
    sleep(3)
    
    try: #verifica se tem o "+" de adicionar mídia, mostrando que o número foi encontrado, e permitindo que mande as faturas para ele
        temnumero = pg.locateOnScreen(r"CobraOi\imagens\adicionar_midia.png", region=(290, 700, 520, 240), confidence=0.7)

        if temnumero: #caso tenha achado o número
            achouNumero = True
            print('- Número {} encontrado! Enviando faturas...'.format(num))
            copy(frases[randint(0, len(frases) - 1)].format(nome))
            sleep(0.5)
            pg.hotkey('ctrl', 'v')
            sleep(1)
            pg.press('enter')
            sleep(1)

            x, y = pg.center(temnumero)
            pg.click(x, y, duration=0.5)
            sleep(0.8)  
            btt = pg.locateOnScreen(r'CobraOi\imagens\doc.png', region=(290, 400, 520, 500), confidence=0.7) #BOTÃO DE DOCUMENTOS, PARA ENVIAR A FATURA
            x, y = pg.center(btt)
            pg.click(x, y, duration=0.5)

    except pg.ImageNotFoundException: #senao
        print('- Número {} não foi encontrado! Passando para o próximo...'.format(num))
        pg.hotkey('ctrl', '1')
        achouNumero = False

def baixaPdf(teclado, nome): #responsável por baixar o pdf da fatura

    global indexf
    sleep(5)
    pg.press('esc')
    
    sleep(1)
    btt = pg.locateOnScreen(r'CobraOi\imagens\baixarPdf.png', region=(1000, 40, 500, 300), confidence = 0.9) #BOTÃO DE BAIXAR PDF
    x, y = pg.center(btt)
    pg.click(x,y, duration=0.5)
    sleep(3.5)
    teclado.type(nome + ' ' + str(indexf))

    sleep(1)
    pg.press("enter")
    sleep(1)
    pg.press('esc')
    sleep(1)

def sairEvoltar(cpfCNPJ, index, opcoes, escolhas): #responsável para quando o algoritmo verifica que é necessario sair e entrar na plataforma   
        
    if escolhas == len(opcoes)-1: #se a escolha de opção atual for a última, ele limpa as opções para outra vez que tiver
        opcoes = []
        
    if len(opcoes) > 1: #se houver mais de uma opção, ele inicia um novo atendimento para ir para a segunda opção
        btt = pg.locateOnScreen(r'CobraOi\imagens\novo_atendimento.png', region=(290, 0, 800, 300), confidence = 0.7) #BOTÃO DE NOVO ATENDIMENTO
        x, y = pg.center(btt)
        pg.click(x,y, duration=0.5)
    
    else: #se as opções acabaram, e podemos passar para o próximo
        
        sleep(2)

        btt = pg.locateOnScreen(r'CobraOi\imagens\herminia_perfil.png', region=(835, 100, 500, 300), confidence=0.7) #BOTÃO DO PERFIL HERMÍNIA
        x, y = pg.center(btt)
        pg.click(x,y, duration=0.5)
        
        sleep(2)
        
        btt = pg.locateOnScreen(r'CobraOi\imagens\portal_selecao.png', region=(835, 100, 500, 300), confidence=0.7) #BOTÃO PARA SAIR DO PORTAL E IR PARA SELEÇÃO
        x, y = pg.center(btt)
        pg.click(x,y, duration=0.5)
        
        sleep(3)
                
        btt = pg.locateOnScreen(r'CobraOi\imagens\opc.png', region=(800, 200, 500, 500), confidence=0.7) #PARA ABRIR AS OPÇÕES VAREJO E EMPRESARIAL
        x, y = pg.center(btt)
        pg.click(x,y, duration=0.5)
        
        sleep(1)
        
        if len(str(cpfCNPJ[index + 1][1])) <= 11: #se for cpf, seleciona a opção "varejo"
            btt = pg.locateOnScreen(r'CobraOi\imagens\varejo.png', region=(300, 200, 700, 400), confidence=0.7) #BOTÃO VAREJO
            x, y = pg.center(btt)
            pg.click(x,y, duration=0.8)
            sleep(1)

        else: #se for cnpj, seleciona a opção "empresarial"
            btt = pg.locateOnScreen(r'CobraOi\imagens\empresarial.png', region=(300, 200, 700, 400), confidence=0.7) #BOTÃO EMPRESARIAL
            x, y = pg.center(btt)
            pg.click(x,y, duration=0.8)
            sleep(1)
            
        btt = pg.locateOnScreen(r'CobraOi\imagens\iniciarrr.png', region=(300, 200, 700, 400), confidence=0.7) #INICIAR
        x, y = pg.center(btt)
        pg.click(x,y, duration=0.8)
        sleep(4)

def temFatura(num, frases, teclado, nome, numeroID): #verifica se tem fatura
    
    sleep(3)
    
    pg.scroll(-1000)

    sleep(1)
    
    fatura = 0
    
    global count_fatura, faturas
    #procura o icone de baixar a fatura
    semPagar = pg.locateOnScreen(r'CobraOi\imagens\baixarfatura.png', region=(970, 250, 250, 600), confidence = 0.75)
    
    if semPagar: #se houver ao menos uma fatura, ele procura por mais
        faturas = list(pg.locateAllOnScreen(r'CobraOi\imagens\baixarfatura.png', region=(970, 250, 250, 600), confidence=0.75))
        count_fatura = len(faturas)
        
        print('- o CPF / CNPJ {} tem um total de {} faturas'.format(numeroID, len(faturas)))

        global indexf

        indexf = 0        

        for quant in faturas: #baixaas faturas, tendo 1 ou mais
        
            indexf+=1
            
            pegarConta = pg.center(quant)
            pg.click(pegarConta, duration=1)
            
            baixaPdf(teclado, nome)
            pg.hotkey("ctrl", "w")  #fecha a aba do pdf    
            sleep(1)
            
        pg.hotkey("ctrl", "2")  #vai pro whatsapp para que possa enviar ao cliente
        msgWhatsapp(num, frases, teclado, nome)
        sleep(4)
        
        if achouNumero: #se tiver número, envia a fatura ao cliente
            selectFaturasMandar(teclado)
            sleep(3)

def avisarFinalizou(tempoExecucao, teclado):
    
    num = '86989030943'
    
    pg.hotkey('ctrl', '2')
    #for num in numeros:
    sleep(1)
    pg.press('esc')
    sleep(1)
    pg.click(351, 116, duration=1.5)
    sleep(1)
    
    teclado.type(num)
    sleep(3)
    pg.press("enter")
    sleep(1)
    
    teclado.type("Automação de cobrança finalizada!")
    pg.press('enter')
    sleep(0.5)
    teclado.type("Tempo de execução: {:.2f} minutos".format(tempoExecucao/60))

def oifibra(numeroID):

    sleep(3)
    pg.scroll(-1000)
    econtaop = pg.locateOnScreen(r'CobraOi\imagens\pesquisaecontas.png', region=(300, 200, 500, 300), confidence=0.9)
    x, y = pg.center(econtaop)
    pg.click(x, y, duration=0.5)
    
    sleep(1)
    econtaop = pg.locateOnScreen(r'CobraOi\imagens\contasfibra.png', region=(300, 400, 500, 300), confidence=0.7)
    x, y = pg.center(econtaop)
    pg.click(x, y, duration=0.5)
    
    sleep(1)
    econtaop = pg.locateOnScreen(r'CobraOi\imagens\cpfecontas.png', region=(600, 300, 500, 300), confidence=0.7)
    x, y = pg.center(econtaop)
    pg.click(x, y+30, duration=0.5)
    
    sleep(1)
    pg.write(numeroID)
    sleep(1)
    
    econtaop = pg.locateOnScreen(r'CobraOi\imagens\enterecontas.png', region=(400, 450, 500, 300), confidence=0.9)
    x, y = pg.center(econtaop)
    pg.click(x, y, duration=0.5)
    
    sleep(2)
    pg.hotkey('ctrl', 'end')

def main():
    
    #arquivo_excel = abrir_arquivo()
    arquivo_excel = r'C:\Users\passo\Downloads\CONTROLE DE PAGAMENTO PARTICULAR.xlsx'
    
    if arquivo_excel:
        
        #planilhaNum = int(input("Digite o número da planilha: "))
        
        #pg.hotkey('alt', 'space', 'n', interval=0.3)
        
        start_time = time() #tempo de execução
        
        planilha = read_excel(arquivo_excel, sheet_name=12, header=None)
        colunas = planilha.iloc[3:, :9]
        
        opcoesIMG = [r'CobraOi\imagens\caso1.png', r'CobraOi\imagens\caso2.png'] #variações de opção   
        opcoes = []
        
        cpfCNPJ = []
        teclado = Controller() #teclado
        
        count_extrato = 0
        count_boleto = 0
        
        frases = [
            "Olá {}, tudo bem? Estou passando para lembrar da(s) sua(s) fatura(s) que não foi(ram) paga(s)!",
            "Olá {}, como vai você? Estou entrando em contato para lembrá-la da(s) sua(s) fatura(s) que ainda não foram pagas.",
            "Oi {}, espero que esteja bem! Só para te lembrar, você tem fatura(s) pendente(s) que precisa(m) ser paga(s)."
        ]
        
        #pg.mouseInfo()
        for index, row in colunas.iterrows():
            #verificando se o valor na segunda coluna tem tamanho de CPF
            
            devendo = 0
            
            verifica = [row[3], row[4], row[5], row[6], row[7], row[8]]
            
            for i, item in enumerate(verifica):
                if notna(item) and item != 'PAGO' and item != 'PAGO ' and item != ' PAGO' and item != 'CANCELADO' and item != 'CANCELADO ' and item != ' CANCELADO' and item != ' CHURN' and item != 'CHURN ' and item != 'CHURN':
                    devendo+=1
            
            if devendo != 0:
                cpfCNPJ.append(row.tolist())

        print(cpfCNPJ)
        
        tempo_total_segundos = len(cpfCNPJ) * 75
        tempo_execucao_minutos = tempo_total_segundos // 60
        
        print("\nTempo de execução aproximado: {} minutos".format(tempo_execucao_minutos))
        
        for index, conta in enumerate(cpfCNPJ):
                
            op = 0
            numeroID = str(conta[1]) #cpf ou cnpj da vez
            #numeroID = str('02877721442')
            num = str(conta[2]) #numero da vez
            #num = str(89981442340)
            nome = str(conta[0]) #nome da vez
            
            if len(num) > 15:
                num = num.split('/')
                num = num[0]
            
            print('\n- Iniciando busca no CPF / CNPJ {}'.format(numeroID))
            
            aa = telaInsereID(numeroID, opcoesIMG, opcoes, op) #insere as informações do cliente
            cliente_cancelado = aa
            print(cliente_cancelado)
            for escolhas, quant in enumerate(opcoes): #loop q percorre todas as opções, preenchendo - as
                print(len(opcoes))
                        
                print('ESCOLHAS:', escolhas)
                op = escolhas
                
                if len(opcoes)>1 and escolhas>0: #se houver mais de uma opção, ele pede para inserir as informações as próximas vezes
                    telaInsereID(numeroID, opcoesIMG, opcoes, op)
                    
                if not cliente_cancelado: #se o cliente não for cancelado
                    
                    sleep(4)
                    fixo() #verificar se é fixo
                    
                    if not fixo(): #se não for fixo
                        
                        sleep(3)
                        telaIniciaAtendimento()
                        
                        sleep(3)
                        telaSegundaEcontas()
                        
                        sleep(6)
                        
                        if simOifibras: #se for Oi Fibra
                            
                            oifibra(numeroID)
                            devendo = devendoEconta()    
                            print(devendo)

                            if devendo: #se tiver fatura pendente no econtas
                                
                                fatecontas = list(pg.locateAllOnScreen(r'CobraOi\imagens\baixarecontas.png', region=(940, 150, 160, 700), confidence=0.9))
                                x, y = pg.center(fatecontas[-1])
                                pg.click(x,y, duration=0.5)
                                
                                sleep(1)
                                baixaPdf(teclado, nome)
                                pg.hotkey('ctrl', 'w')
                                sleep(1)
                                pg.hotkey("ctrl", "2")  #vai pro whatsapp para que possa enviar ao cliente
                                msgWhatsapp(num, frases, teclado, nome)
                                sleep(4)
                                
                                if achouNumero: #se tiver número, envia a fatura ao cliente
                                    selectFaturasMandar(teclado)
                                    sleep(3)
                                
                                sairEvoltar(cpfCNPJ, index, opcoes, escolhas)
                                
                            else: # se não tiver fatura pendente no econtas
                                print('sim')
                                sairEvoltar(cpfCNPJ, index, opcoes, escolhas)
                                
                        else: #se não for fatura Oi Fibra
                            
                            try:        
                                temFatura(num, frases, teclado, nome, numeroID)
                                sairEvoltar(cpfCNPJ, index, opcoes, escolhas)
                                
                            except pg.ImageNotFoundException: #se caso não seja encontrado nenhuma fatura pendente no normal
                                sairEvoltar(cpfCNPJ, index, opcoes, escolhas)
                                
                    else: #se o serviço for Fixo  
                        sairEvoltar(cpfCNPJ, index, opcoes, escolhas)
                elif cliente_cancelado == True: #se o cliente for cancelado
                    print('teste')
                    sairEvoltar(cpfCNPJ, index, opcoes, escolhas)
                            
        print(cliente_cancelado)
        
        #finaliza o tempo                    
        end_time = time()
        execution_time = end_time - start_time
        
        avisarFinalizou(execution_time, teclado)
        
    else: #se nenhum arquivo for selecionado
        print("- Nenhum arquivo selecionado. Encerrando o programa.")

if __name__ == "__main__":
    
    try:
        main()
        
    except:
        teclado = Controller()
                
        num = '86989030943'
        
        pg.hotkey('ctrl', '2')
        #for num in numeros:
        sleep(1)
        pg.press('esc')
        sleep(1)
        pg.click(351, 116, duration=1.5)
        sleep(1)
        
        teclado.type(num)
        sleep(3)
        pg.press("enter")
        sleep(1)
        
        teclado.type("ACONTECEU UM ERRO: \n\n\n\n")
        sleep(1)
        pg.press('enter')
        
        error_message = traceback.format_exc()
        
        copy(error_message)
        sleep(1)
        pg.hotkey('ctrl', 'v')
        sleep(2)
        pg.press('enter')