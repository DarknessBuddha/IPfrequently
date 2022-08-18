import time
import sys
import os
import pyautogui
import win32gui
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def get_highlighted_coords(*, color=(255, 150, 50)):
    s = pyautogui.screenshot()
    for x in range(s.width):
        for y in range(s.height):
            if s.getpixel((x, y)) == color:
                return x, y
    return -1, -1


def move_to_highlighted(*, color=(255, 150, 50), offset_x=0, offset_y=0):
    x, y = get_highlighted_coords(color=color)
    pyautogui.moveTo(x+offset_x, y+offset_y)


def click_on_highlighted(*, offset_x=0, offset_y=0):
    x, y = get_highlighted_coords()
    if (x, y) == (-1, -1):
        raise Exception('Error highlight not found')
    pyautogui.click(x + offset_x, y + offset_y)


def find_by_text(text: str):
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(.1)
    pyautogui.write(text)


def click_on_link(text: str):
    find_by_text(text)
    pyautogui.hotkey('ctrl', 'enter')


def click_on_input_box_lazy(text: str):
    find_by_text(text)
    time.sleep(.2)
    click_on_highlighted()
    time.sleep(.2)
    pyautogui.press('tab')

def click_on_input_box_fast(text: str):
    find_by_text(text)
    time.sleep(.2)
    pyautogui.hotkey('ctrl', 'enter')
    time.sleep(.2)
    pyautogui.press('tab')

def click_on_input_box_greedy(text: str, search_direction='right', *, cache=[-1, -1]):
    """search_direction: right, down"""
    distance = 25
    find_by_text(text)
    move_to_highlighted()

    if search_direction == 'down':
        if cache[1] != -1:
            pyautogui.move(0, cache[1])
            if win32gui.GetCursorInfo()[1] != 65541:
                pyautogui.click()
                return

        move_to_highlighted()
        while win32gui.GetCursorInfo()[1] == 65541:
            pyautogui.move(0, distance)
            cache[1] += distance

    else:

        if cache[0] != -1:
            pyautogui.moveTo(cache[0], pyautogui.position()[1])
            if win32gui.GetCursorInfo()[1] != 65541:
                pyautogui.click()
                return

        move_to_highlighted()
        while win32gui.GetCursorInfo()[1] == 65541:
            pyautogui.move(distance, 0)
        cache[0] = pyautogui.position()[0]
    pyautogui.click()


def get_file_name():
    Tk().withdraw()
    return askopenfilename()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)