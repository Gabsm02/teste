import pyautogui as py
import time


x,y = py.position()
print(x,y)

py.moveTo(1035,1036, duration=3)
py.click()
# Clicar no chrome
py.moveTo(328,147, duration=3)
py.click()
# Entrou no maestro
time.sleep(5)
py.moveTo(943,698, duration=3)
py.click()
# Fez login
time.sleep(7)
py.moveTo(945,491, duration=2)
py.click()
# Clicou na Seta
py.moveTo(770,547, duration=3)
py.click()
time.sleep(5)
py.moveTo(995,828, duration=3)
py.click()
# Selecionou "1 dia"
# time.sleep(10)
# py.moveTo(660,933, duration=3)
# py.click()
# # Selecionou "NE"
time.sleep(2)
py.scroll(-950)
py.moveTo(1782,758, duration=3)
py.click()
# Clickou em download
py.moveTo(1755,830, duration=3)
py.click()
# # Selecionou excel
py.moveTo(1542,213, duration=2)
py.click()
# time.sleep(20)

