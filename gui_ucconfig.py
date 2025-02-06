import PySimpleGUI as sg
import pprint
import json
import re
import os
from itertools import cycle
import base64
from axl import axl
from ris import ris
from ucxnobjects import *
import zeep.helpers
import colorama
from colorama import Fore, Back, Style

DEFAULT_SETTINGS = {'cucm_user': 'AXLpy', 'cucm_pwd': None , 'ucxn_user': 'RESTpy', 'ucxn_pwd' : 'ABC124'}
# "Map" from the settings dictionary keys to the window's element keys
SETTINGS_KEYS_TO_ELEMENT_KEYS = {'CUCM.USERNAME': '-CUCM USER-', 'CUCM.PASSWORD': '-CUCM PWD-' , 'UCXN.USERNAME': '-UCXN USER-', 'UCXN.PASSWORD': '-UCXN PWD-'}

def get_json_data(filename):
    jf = open(filename, 'r')
    ucfg= json.load(jf)
    print(Fore.CYAN + 'File -',filename)
    ucfg['CUCM']['PASSWORD'] = xor_crypt_string(ucfg['CUCM']['PASSWORD'], decode=True)
    ucfg['UCXN']['PASSWORD'] = xor_crypt_string(ucfg['UCXN']['PASSWORD'], decode=True)
    pprint.pprint(ucfg)
    jf.close()
    return ucfg


def xor_crypt_string(data, key='G', encode=False, decode=False):
    if decode:
        b64_data_b = data.encode("ascii")
        data_b = base64.b64decode(b64_data_b)
        data = data_b.decode('ascii')

    xored = ''.join(chr(ord(x) ^ ord(y)) for (x, y) in zip(data, cycle(key)))

    if encode:
        xored_b = xored.encode('ascii')
        b64_xored_b = base64.b64encode(xored_b)
        return b64_xored_b.decode('ascii').strip()
    return xored

def TextLabel(text): return sg.Text(text+' : ', justification='r', size=(15,1))

def save_settings(filename, values):
    if values:      # if there are stuff specified by another window, fill in those values
        ucfg = get_json_data(filename)
        valuesgood = True
        for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:  # update window with the values read from settings file
            try:
                lkey = key.split('.')
                if lkey[1] != 'PASSWORD':
                    ucfg[lkey[0]][lkey[1]] = values[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]]
                else:
                    ucfg[lkey[0]][lkey[1]] = xor_crypt_string(values[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]], encode=True)
            except Exception as e:
                print(f'Problem updating settings from window values. Key = {key}')
                valuesgood = False
        if valuesgood:
            with open(filename, 'w') as json_out:
                print(json.dumps(ucfg, indent = 2),file=json_out)
            print(Fore.LIGHTGREEN_EX + 'File -', filename)
            pprint.pprint(ucfg)

def process_axl_test(ucfg):
    cucm = ucfg['CUCM']['HOST']
    cucm_username = ucfg['CUCM']['USERNAME']
    cucm_password = ucfg['CUCM']['PASSWORD']
    cucm_version = '12.5'

    ucm = axl(username=cucm_username, password=cucm_password, cucm=cucm, cucm_version=cucm_version)

    cucm_resp = ucm.get_call_manager_group('Default')
    cucm_in_dict = zeep.helpers.serialize_object(cucm_resp)
    if isinstance(cucm_in_dict,dict):
        cucm_test = 'CUCM Test Successful'
    else:
        cucm_test = 'CUCM Test NOT Successful'
    return(cucm_test)

def process_rest_test(ucfg):
    ucxn = ucfg['UCXN']['HOST']
    ucxn_username = ucfg['UCXN']['USERNAME']
    ucxn_password = ucfg['UCXN']['PASSWORD']
    ucxn_version = '12.5'

    ucm = callhandlers(username=ucxn_username, password=ucxn_password, ucxn=ucxn, ucxn_version=ucxn_version,apitype='callhandlers')
    chobj = f"DisplayName%20startswith%20Goodbye"
    ucxn_resp = ucm.ucxnlistwq(chobj)

    if isinstance(ucxn_resp,dict) and 'error' not in ucxn_resp:
        ucxn_test = 'UCXN Test Successful'
    else:
        ucxn_test = 'UCXN Test NOT Successful'
    return(ucxn_test)

def test_settings(filename, values):
    file = re.split('ucconfig-',filename)
    if len(file) > 1:
        cluster = re.split(file[1],'.')
    else:
        cluster = 'na'

    ucfg = get_json_data(filename)
    cucm_test = process_axl_test(ucfg)
    ucxn_test = process_rest_test(ucfg)
    uc_test = f'{cucm_test} - {ucxn_test}'
    return(uc_test)
def cntr(elems):
        return[sg.Stretch(), *elems,sg.Stretch()]

def main():
    colorama.init()
    sg.ChangeLookAndFeel('DarkGreen'),
    layout = [[sg.Text('Please Select Input File:', size=(20, 1)), sg.Input(), sg.FileBrowse(file_types=(("UCconfig Files", "ucconfig*.json"),))],
              [sg.ReadButton('Open', size = (13,1))],
              [sg.Text('_'  * 80)],
              [sg.Text('CUCM')],
              [TextLabel('Username'),sg.InputText(key='-CUCM USER-')],
              [TextLabel('Password'), sg.InputText(key='-CUCM PWD-')],
              [sg.Text(' '  * 80)],
              [sg.Text('UCXN')],
              [TextLabel('Username'), sg.InputText(key='-UCXN USER-')],
              [TextLabel('Password'), sg.InputText(key='-UCXN PWD-')],
              [sg.Text(' '  * 80)],
              [sg.Text('', key='-MESSAGE-',text_color='yellow',background_color='DarkGreen')],
              [sg.Submit('Save'), sg.Submit('Test'), sg.Cancel()]]

    window = sg.Window('Config Changes', layout, default_element_size=(40, 1), grab_anywhere=False)
    filename = 'ucconfig.json'
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == 'Save':
            print(Fore.WHITE, '---')
            pprint.pprint(values)
            save_settings(filename, values)
            window['-MESSAGE-'].Update('Save Successful')
        elif event == 'Test':
            uc_test = test_settings(filename, values)
            window['-MESSAGE-'].Update(uc_test)
        elif event == 'Cancel':
            break
        elif event == 'Open':
            if values[0] != '':
                filename = values[0]
                jdict = get_json_data(filename)
                window['-CUCM USER-'].Update(jdict['CUCM']['USERNAME'])
                window['-CUCM PWD-'].Update(jdict['CUCM']['PASSWORD'])
                window['-UCXN USER-'].Update(jdict['UCXN']['USERNAME'])
                window['-UCXN PWD-'].Update(jdict['UCXN']['PASSWORD'])
            else:
                filename = 'ucconfig.json'
                sg.popup_quick_message(f'No config file found... please browse to one then press Open', keep_on_top=True,
                                   background_color='orange', text_color='blue')
        else:
            break

    window.close()
    #pprint.pprint(values)

main()