import pyautogui as pg

semPagar = pg.locateOnScreen(r'CobraOi\imagens\baixarfatura.png', region=(970, 250, 250, 600), confidence = 0.75)

x,y=pg.center(semPagar)
pg.moveTo(x,y)
