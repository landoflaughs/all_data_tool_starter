# coding=utf-8
import sys
from time import sleep

from win32con import WM_INPUTLANGCHANGEREQUEST
from string_resource import key_input

import win32api
import win32gui
from pykeyboard import PyKeyboard
import os


def change_language(lang="en"):
    LID = {
        "ch": 0x0804,
        "en": 0x0409
    }
    hwnd = win32gui.GetForegroundWindow()
    language = LID[lang]
    result = win32api.SendMessage(
        hwnd,
        0,
        WM_INPUTLANGCHANGEREQUEST,
        language
    )
    if result == 0:
        return True


# handle upper word
def type_keys(content):
    contents = list(content)
    for i in range(len(contents)):
        if 'A' <= contents[i] <= 'Z':
            k.tap_key(k.caps_lock_key)
            lowercont = str(contents[i].swapcase())
            # print(lowercont)
            key_input(lowercont)
            k.tap_key(k.caps_lock_key)
        if 'a' <= contents[i] <= 'z':
            key_input(contents[i])
        if contents[i] in ['_', '.', '/', '#', ' ', '\\', ':']:
            k.type_string(contents[i])
        if '0' <= contents[i] <= '9':
            key_input(contents[i])



def navigate_to_txtfolder(path):
    if path == '':
        print('no config txt path, use data_all_tools path')
        return
    if path != '':
        if '#' in path:
            print('no txt path set, use data_all_tools path')
            return
        if '#' not in path:
            type_keys(path) # not support CH
            sleep(0.1)
            k.tap_key(k.enter_key)  # click enter to preconfigured folder
            print("changed to path: ")
            print(path)
            sleep(1)


def input_user_psw(server_ip, database_name, database_pwd, database_model, txt_path):
    k.tap_key(k.tab_key)
    k.tap_key(k.tab_key)
    key_input(server_ip)
    sleep(0.2)
    k.tap_key(k.tab_key)

    type_keys(database_name)
    sleep(0.2)
    k.tap_key(k.tab_key)

    type_keys(database_pwd)
    sleep(0.2)
    k.tap_key(k.tab_key)

    type_keys(database_model)

    sleep(0.2)
    k.tap_key(k.enter_key)
    sleep(1)
    k.tap_key(k.enter_key)  # enter tool
    sleep(0.2)
    k.tap_key(k.tab_key)  # move to preview button


    sleep(0.2)
    k.tap_key(k.enter_key)  # open preview selections
    sleep(0.2)
    navigate_to_txtfolder(txt_path)


def start_exe(tool_path):
    os.startfile(tool_path)



if __name__ == '__main__':
    print ('please alter your keyboard init mode to  "EN" first (if you use QQ pinyin, SouGou pinyin,etc.) \n\n')
    number = input("enter the number of tools you want to activate: ")

    txt_path = ''

    with open(r"datas\config.ini", "r") as  f:
        login_datas = f.readlines()
        print(login_datas)
        ip = login_datas[0].replace('\n', '').strip(' ')
        name = login_datas[1].replace('\n', '').strip(' ')
        pwd = login_datas[2].replace('\n', '').strip(' ')
        model = login_datas[3].replace('\n', '').strip(' ')
        to_path = login_datas[4].replace('\n', '').strip(' ')

    if os.path.exists('txt_path.ini'):
        with open(r"txt_path.ini", "r") as f:
            txt_paths = f.readlines()
            txt_path = txt_paths[0].replace('\n', '').strip(' ')

    for i in range(number):
        start_exe(to_path)
        sleep(1)
        change_language()
        k = PyKeyboard()
        input_user_psw(ip, name, pwd, model, txt_path)
