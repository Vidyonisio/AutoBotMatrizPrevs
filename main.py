import pyautogui
import time
import pygetwindow as gw
import ctypes
from ctypes import wintypes
from collections import namedtuple

"Automatiza o preenchimento da curva de carga nas boletas do Thunder via Bot"

print("IMPORTANTE: É necessário estar com a planilha com os prevs aberta e as pastas destino, a célula da ENA S selecionada")

se = 60
s = 60
s_txt = str(s)+"%"
se_txt = str(se)+"%"
dest = input('Digite o Título da Pasta Destino: ')
dest = str(dest)
fon_nomes = input('Digite o Título da Pasta De onde virão os nomes: ')
# pyautogui.getWindowsAt(15528,10)[0].activate()
time.sleep(.1)

def ciclo(s_txt, dest, Fon_nomes):
    gw.getWindowsWithTitle('Vazoes_Prevs')[0].activate()
    gw.getWindowsWithTitle('Vazoes_Prevs')[0].maximize()
    time.sleep(.4)
    pyautogui.press('up')
    time.sleep(.4)
    pyautogui.press('f2')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'a')
    time.sleep(.4)
    pyautogui.write(s_txt)
    time.sleep(.4)
    pyautogui.press('enter')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'm')
    time.sleep(.4)
    gw.getWindowsWithTitle('Prevs-Pasta')[0].activate()
    gw.getWindowsWithTitle('Prevs-Pasta')[0].maximize()
    time.sleep(.4)
    pyautogui.press('up')
    time.sleep(.4)
    pyautogui.press('up')
    time.sleep(.4)
    pyautogui.press('up')
    time.sleep(.4)
    pyautogui.press('enter')
    time.sleep(.4)
    pyautogui.press('down')
    time.sleep(.4)
    pyautogui.press('down')
    time.sleep(.4)
    pyautogui.press('down')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'c')
    gw.getWindowsWithTitle(dest)[0].activate()
    pyautogui.hotkey('ctrlleft', 'v')
    gw.getWindowsWithTitle(Fon_nomes)[0].activate()
    pyautogui.press('f2')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'a')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'c')
    time.sleep(.4)
    pyautogui.press('esc')
    time.sleep(.4)
    pyautogui.press('down')
    time.sleep(.4)
    gw.getWindowsWithTitle(dest)[0].activate()
    time.sleep(.4)
    pyautogui.press('f2')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'a')
    time.sleep(.4)
    pyautogui.hotkey('ctrlleft', 'v')
    time.sleep(.4)
    pyautogui.press('enter')
    time.sleep(.4)
    pyautogui.press('enter')




for j in range(6):
    ciclo(s_txt, dest, fon_nomes)
    s += 10
    s_txt = str(s)+"%"
    time.sleep(.2)
    j+=1

# for i in range(5):   #Peenchimento do primeiro ano até o final
#     for j in range(6):
#         s += 10
#     se += 10


# Daqui pra baixo código pra seber o nome das pastas

# user32 = ctypes.WinDLL('user32', use_last_error=True)
# def check_zero(result, func, args):
#     if not result:
#         err = ctypes.get_last_error()
#         if err:
#             raise ctypes.WinError(err)
#     return args
#
# if not hasattr(wintypes, 'LPDWORD'): # PY2
#     wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)
#
# WindowInfo = namedtuple('WindowInfo', 'pid title')
#
# WNDENUMPROC = ctypes.WINFUNCTYPE(
#     wintypes.BOOL,
#     wintypes.HWND,    # _In_ hWnd
#     wintypes.LPARAM,) # _In_ lParam
#
# user32.EnumWindows.errcheck = check_zero
# user32.EnumWindows.argtypes = (
#    WNDENUMPROC,      # _In_ lpEnumFunc
#    wintypes.LPARAM,) # _In_ lParam
#
# user32.IsWindowVisible.argtypes = (
#     wintypes.HWND,) # _In_ hWnd
#
# user32.GetWindowThreadProcessId.restype = wintypes.DWORD
# user32.GetWindowThreadProcessId.argtypes = (
#   wintypes.HWND,     # _In_      hWnd
#   wintypes.LPDWORD,) # _Out_opt_ lpdwProcessId
#
# user32.GetWindowTextLengthW.errcheck = check_zero
# user32.GetWindowTextLengthW.argtypes = (
#    wintypes.HWND,) # _In_ hWnd
#
# user32.GetWindowTextW.errcheck = check_zero
# user32.GetWindowTextW.argtypes = (
#     wintypes.HWND,   # _In_  hWnd
#     wintypes.LPWSTR, # _Out_ lpString
#     ctypes.c_int,)   # _In_  nMaxCount
# def list_windows():
#     '''Return a sorted list of visible windows.'''
#     result = []
#     @WNDENUMPROC
#     def enum_proc(hWnd, lParam):
#         if user32.IsWindowVisible(hWnd):
#             pid = wintypes.DWORD()
#             tid = user32.GetWindowThreadProcessId(
#                         hWnd, ctypes.byref(pid))
#             length = user32.GetWindowTextLengthW(hWnd) + 1
#             title = ctypes.create_unicode_buffer(length)
#             user32.GetWindowTextW(hWnd, title, length)
#             result.append(WindowInfo(pid.value, title.value))
#         return True
#     user32.EnumWindows(enum_proc, 0)
#     return sorted(result)
# print(list_windows())
#
# def ciclo_s_inic(s):
#     gw.getWindowsWithTitle('Vazoes_Prevs')[0].activate()
#     time.sleep(.2)
#     pyautogui.press('f2')
#     time.sleep(.2)
#     pyautogui.hotkey('ctrlleft', 'a')
#     time.sleep(.2)
#     pyautogui.write(s_txt)
#     time.sleep(.2)
#     pyautogui.press('enter')
#     time.sleep(.2)
#     pyautogui.hotkey('ctrlleft', 'm')
#     gw.getWindowsWithTitle('Prevs')[0].activate()
#     pyautogui.press('up')
#     pyautogui.press('up')
#     pyautogui.press('up')
#     pyautogui.press('enter')
#     time.sleep(.5)
#     pyautogui.press('down')
#     pyautogui.press('down')
#     pyautogui.press('down')
#     pyautogui.hotkey('ctrlleft', 'c')
#     pyautogui.keyDown('alt')
#     time.sleep(.2)
#     pyautogui.press('tab')
#     time.sleep(.2)
#     pyautogui.press('tab')
#     time.sleep(.2)
#     pyautogui.press('tab')
#     time.sleep(.2)
#     pyautogui.keyUp('alt')
#     pyautogui.hotkey('ctrlleft', 'v')



