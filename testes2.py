import pyautogui as pg

opcoesIMG = [r'CobraOi\imagens\caso1.png', r'CobraOi\imagens\caso2.png']
opcoes = []

for ft in opcoesIMG:
    try:
        casos = pg.locateAllOnScreen(ft, region=(480, 430, 490, 240), confidence=0.7)
        
        if casos:

            for index, caso in enumerate(casos):
                opcoes.append(pg.center(caso))
                pg.moveTo(opcoes[index][0], opcoes[index][1])
                
            print(opcoes)
    except pg.ImageNotFoundException:

        continue
    finally:
        continue 

