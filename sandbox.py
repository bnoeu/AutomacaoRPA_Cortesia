import time
import pyautogui as bot
from utils.funcoes import ahk as ahk
from utils.funcoes import procura_imagem


ahk.win_activate('TopCompras', title_match_mode= 2)
ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)

if procura_imagem('imagens/img_topcon/cadastramento_bruno.png', continuar_exec= True, msg_confianca= True):
    bot.alert("Encontrou!")
#bot.click(procura_imagem('imagens/img_topcon/cadastramento_bruno.png', continuar_exec= True, msg_confianca= True))