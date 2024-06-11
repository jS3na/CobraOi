import pyautogui as pg

btt = pg.locateOnScreen(r'.\imagens\doc.png', region=(290, 400, 520, 500), confidence=0.7) #BOTÃO DE DOCUMENTOS, PARA ENVIAR A FATURA
x, y = pg.center(btt)
pg.click(x, y, duration=0.5)

