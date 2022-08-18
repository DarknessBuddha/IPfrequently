import time
import pyautogui
import pandas
import utils
import sys
from pandas import DataFrame


def run_bot(file, sheet_name, ranges, existing_name, name, mac):
    if file == '':
        sys.exit()
    APs = pandas.read_excel(file, sheet_name=sheet_name if not sheet_name.isnumeric() else int(sheet_name))
    print(name, mac)
    # execution
    for start, end in ranges:

        if start <= end:
            selected_APs: DataFrame = APs.iloc[start - 2:end - 1]
        else:
            selected_APs: DataFrame = APs.iloc[start - 2:(end - 3 if end - 3 >= 0 else None):-1]

        for index, AP in selected_APs.iterrows():
            try:
                if AP[mac] == '' or AP[mac] is None or AP[mac] == ':::::':
                    continue

                # click edit
                utils.find_by_text(AP['IP Address'])
                time.sleep(.5)
                pyautogui.hotkey('ctrl', 'enter')
                time.sleep(.2)
                pyautogui.hotkey('shift', 'tab')
                time.sleep(.4)
                pyautogui.hotkey('shift', 'tab')
                time.sleep(.1)
                pyautogui.press('enter')
                time.sleep(1)

                # fill description
                if existing_name:
                    utils.find_by_text(AP[existing_name])
                    pyautogui.hotkey('ctrl', 'enter')
                else:
                    utils.click_on_input_box_fast('description')
                    pyautogui.hotkey('ctrl', 'a')
                time.sleep(.1)
                pyautogui.write(AP[name])

                # fill hardware address
                utils.click_on_input_box_fast('hardware address')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.write(AP[mac])

                # save changes
                utils.find_by_text('save changes')
                time.sleep(.1)
                pyautogui.hotkey('ctrl', 'enter')
                time.sleep(.5)

                # go back to search
                utils.find_by_text('search for address')
                time.sleep(.1)
                pyautogui.hotkey('ctrl', 'enter')
                time.sleep(1)

            except Exception as e:
                print(e)
                break
    print('finished')
