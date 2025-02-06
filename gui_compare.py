import PySimpleGUI as sg
import os
import pprint
import ast
import pandas as pd
import numpy as np
from pandasgui import show
import colorama
from colorama import Fore, Back, Style

def report_diff(x):
    """Function to use with groupby.apply to highlight value changes."""
    return x[0] if x[0] == x[1] or pd.isna(x).all() else f'{x[0]} ---> {x[1]}'

def highlight_diff(x):
    """Function to use with style.applymap to highlight value changes."""
    if isinstance(x,str):
        if "--->" in x:
            return 'background-color: orange'
            #return 'background-color: %s' % 'orange'
    return ''

def strip(x):
    """Function to use with applymap to strip whitespaces in a dataframe."""
    return x.strip() if isinstance(x, str) else x


def diff_pd(old_df, new_df, idx_col):
    """Identify differences between two pandas DataFrames using a key column.
    Key column is assumed to have only unique data
    (like a database unique id index)
    Args:
        old_df (pd.DataFrame): first dataframe
        new_df (pd.DataFrame): second dataframe
        idx_col (str|list(str)): column name(s) of the index,
          needs to be present in both DataFrames
    """
    # setting the column name as index for fast operations
    out_data = {
        'error': '',
    }
    try:
        old_df = old_df.set_index(idx_col)
    except KeyError as errmsg:
        out_data['error'] = str(errmsg) + ' in file 1'
        return (out_data)
    try:
        new_df = new_df.set_index(idx_col)
    except KeyError as errmsg:
        out_data['error'] = str(errmsg) + ' in file 2'
        return (out_data)

    # get the added and removed rows
    old_keys = old_df.index
    new_keys = new_df.index
    if isinstance(old_keys, pd.MultiIndex):
        removed_keys = old_keys.difference(new_keys)
        added_keys = new_keys.difference(old_keys)
    else:
        removed_keys = np.setdiff1d(old_keys, new_keys)
        added_keys = np.setdiff1d(new_keys, old_keys)
    out_data = {
        'error': '',
        'removed row': old_df.loc[removed_keys],
        'added row': new_df.loc[added_keys]
    }
    # focusing on common data of both dataframes
    common_keys = np.intersect1d(old_keys, new_keys, assume_unique=True)
    common_columns = np.intersect1d(
        old_df.columns, new_df.columns, assume_unique=True
    )
    new_common = new_df.loc[common_keys, common_columns].applymap(strip)
    old_common = old_df.loc[common_keys, common_columns].applymap(strip)
    # get the changed rows keys by dropping identical rows
    # (indexes are ignored, so we'll reset them)
    common_data = pd.concat(
        [old_common.reset_index(), new_common.reset_index()], sort=True
    )
    changed_keys = common_data.drop_duplicates(keep=False)[idx_col]
    if isinstance(changed_keys, pd.Series):
        changed_keys = changed_keys.unique()
    else:
        changed_keys = changed_keys.drop_duplicates().set_index(idx_col).index
    # combining the changed rows via multi level columns
    df_all_changes = pd.concat(
        [old_common.loc[changed_keys], new_common.loc[changed_keys]],
        axis='columns',
        keys=['old', 'new']
    ).swaplevel(axis='columns')
    # using report_diff to merge the changes in a single cell with "-->"
    df_changed = df_all_changes.groupby(level=0, axis=1).apply(
        lambda frame: frame.apply(report_diff, axis=1))

    styleddf = (df_changed.style.applymap(highlight_diff))
    #out_data['changed row'] = df_changed
    out_data['changed row'] = styleddf

    # get the added and removed columns
    old_col = list(old_df.columns)
    new_col = list(new_df.columns)
    removed_col = np.setdiff1d(old_col, new_col)
    added_col = np.setdiff1d(new_col, old_col)

    if(removed_col.size > 0):
        old_diff = old_df.loc[common_keys, removed_col].applymap(strip)
        out_data['removed col'] = old_diff

    if(added_col.size > 0):
        new_diff = new_df.loc[common_keys, added_col].applymap(strip)
        out_data['added col'] = new_diff

    return out_data

def compare_excel(path1, sheet1, path2, sheet2, out_path, index_col_name, **kwargs):
    p1 = path1.split('.')
    p2 = path2.split('.')
    pdir = os.path.dirname(path1) + '/'
    if (p1[1] == 'csv'):
        try:
            old_df = pd.read_csv(path1, dtype = str, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        old_df = old_df.fillna('')
    else:
        try:
            old_df = pd.read_excel(path1, sheet_name=sheet1, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        old_df = old_df.fillna('')

    if (p2[1] == 'csv'):
        try:
            new_df = pd.read_csv(path2, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        new_df = new_df.fillna('')
    else:
        try:
            new_df = pd.read_excel(path2, sheet_name=sheet2, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        new_df = new_df.fillna('')

    diff = diff_pd(old_df, new_df, index_col_name)
    if diff['error'] == '':
        dir_path = pdir + out_path
        with pd.ExcelWriter(dir_path) as writer:
            for sname, data in diff.items():
                if sname != 'error':
                    data.to_excel(writer, engine='openpyxl',sheet_name=sname)
        output_msg = f"Differences saved in {dir_path}"
        print(output_msg)
        return(output_msg)
    else:
        return(diff['error'])

def display_excel(
        path1, sheet1, **kwargs
):
    p1 = path1.split('.')
    pdir = os.path.dirname(path1) + '/'
    if (p1[1] == 'csv'):
        try:
            old_df = pd.read_csv(path1, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        old_df = old_df.fillna('')
    else:
        try:
            old_df = pd.read_excel(path1, sheet_name=sheet1, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        old_df = old_df.fillna('')

    show(old_df)
    output_msg = f"Opened {path1}"
    print(output_msg)
    return(output_msg)

def merge_excel(
        path1, sheet1, path2, sheet2, out_path, index_col_name, **kwargs
):
    p1 = path1.split('.')
    p2 = path2.split('.')
    pdir = os.path.dirname(path1) + '/'
    if (p1[1] == 'csv'):
        try:
            old_df = pd.read_csv(path1, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        old_df = old_df.fillna('')
    else:
        try:
            old_df = pd.read_excel(path1, sheet_name=sheet1, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        old_df = old_df.fillna('')

    if (p2[1] == 'csv'):
        try:
            new_df = pd.read_csv(path2, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        new_df = new_df.fillna('')
    else:
        try:
            new_df = pd.read_excel(path2, sheet_name=sheet2, **kwargs)
        except FileNotFoundError as errmsg:
            return (errmsg)
        except ValueError as errmsg:
            return (errmsg)
        new_df = new_df.fillna('')

    combpd= pd.merge(old_df, new_df, on=index_col_name, how='outer')
    out_data = {
        'merged rows': combpd
    }
    dir_path = pdir + out_path
    with pd.ExcelWriter(dir_path) as writer:
        for sname, data in out_data.items():
            data.to_excel(writer, engine='openpyxl',sheet_name=sname)
    output_msg = f"Combined saved in {dir_path}"
    print(output_msg)
    return(output_msg)


def main():
    colorama.init()
    sg.ChangeLookAndFeel('LightBlue'),
    layout = [[sg.Text('Enter File 1:', size=(20, 1)), sg.Input(), sg.FileBrowse(file_types=(("Excel Files", "*.*"),))],
              [sg.Text('File 1 Sheet Name:', size=(15, 1)), sg.Input()],
              [sg.Text('Enter File 2:', size=(20, 1)), sg.Input(), sg.FileBrowse(file_types=(("Excel Files", "*.*"),))],
              [sg.Text('File 2 Sheet Name:', size=(15, 1)), sg.Input()],
              [sg.Text('_'  * 80)],
              [sg.Text('Key Column:', size=(15, 1)), sg.Input()],
              [sg.Text('', key='-MESSAGE-',text_color='yellow',background_color='DarkGreen')],
              [sg.Submit('Display'), sg.Submit('Compare'), sg.Submit('Merge'), sg.Cancel()]]

    window = sg.Window('Excel Compare', layout, default_element_size=(40, 1), grab_anywhere=False)
    opt = {}
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == 'Save':
            print(Fore.WHITE, '---')
            pprint.pprint(values)
            #save_settings(filename, values)
            window['-MESSAGE-'].Update('Save Successful')
        elif event == 'Test':
            #uc_test = test_settings(filename, values)
            uc_test = ' '
            window['-MESSAGE-'].Update(uc_test)
        elif event == 'Cancel':
            break
        elif event == 'Display':
            if values[0] != '':
                path1 = values[0]
                if values[1] != '':
                    sheet1 = values[1]
                else:
                    sheet1 = 'sheet1'
                output_msg = display_excel(path1, sheet1)
                window['-MESSAGE-'].Update(output_msg)
        elif event == 'Compare':
            if values[0] != '':
                path1 = values[0]
                if values[1] != '':
                    sheet1 = values[1]
                else:
                    sheet1 = 'sheet1'
                path2 = values[2]
                if values[3] != '':
                    sheet2 = values[3]
                else:
                    sheet2 = 'sheet1'
                output_path = "compared.xlsx"
                key_column = ast.literal_eval(values[4])
                output_msg = compare_excel(path1, sheet1, path2, sheet2, output_path, key_column)
                window['-MESSAGE-'].Update(output_msg)
            else:
                sg.popup_quick_message(f'File NOT found... please browse to one then press action', keep_on_top=True,
                                   background_color='orange', text_color='blue')
        elif event == 'Merge':
            if values[0] != '':
                path1 = values[0]
                sheet1 = values[1]
                path2 = values[2]
                sheet2 = values[3]
                output_path = "merged.xlsx"
                key_column = values[4]
                output_msg = merge_excel(path1, sheet1, path2, sheet2, output_path,
                                           key_column)
                window['-MESSAGE-'].Update(output_msg)
            else:
                sg.popup_quick_message(f'File NOT found... please browse to one then press action', keep_on_top=True,
                                   background_color='orange', text_color='blue')
        else:
            break

    window.close()
    #pprint.pprint(values)

main()