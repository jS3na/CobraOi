import cv2
import numpy as np
import pyautogui as pg

pg.mouseInfo()

count_escolha = 0

def devendoEconta():
    global count_escolha, fotoEscolhas
    
    fotoEscolhas = pg.screenshot(r'imagens/escolhaaaaaaaa.png', region=(460, 500, 520, 100))

    # Carregar a imagem a ser procurada
    escolha_template = cv2.imread(r'imagens/escolha.png', cv2.IMREAD_GRAYSCALE)

    # Carregar a imagem na qual será feita a busca
    tela = cv2.imread(r'imagens/escolhaaaaaaaa.png', cv2.IMREAD_GRAYSCALE)

    # Executar a correspondência de padrões
    result = cv2.matchTemplate(tela, escolha_template, cv2.TM_CCOEFF_NORMED)

    # Definir um limite de confiança
    threshold = 0.44
    loc = np.where(result >= threshold)

    # Contar o número de correspondências encontradas
    count_escolha = len(loc[0])

    if count_escolha > 0:
        print("Escolhas: {}".format(count_escolha))

        # Desenhar retângulos nas correspondências encontradas
        for pt in zip(*loc[::-1]):
            cv2.rectangle(tela, pt, (pt[0] + escolha_template.shape[1], pt[1] + escolha_template.shape[0]), (0, 255, 255), 2)

        # Mostrar imagem com as correspondências detectadas (opcional)
        cv2.imshow('Correspondências', tela)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

        return True
    else:
        return False

print(devendoEconta())
