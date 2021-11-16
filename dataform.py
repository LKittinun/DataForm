import PySimpleGUI as sg
import pandas as pd
import numpy as np
import os
import sys
import logging
from datetime import date
from pathlib import Path

class error_handler:

    """[summary]
    Provide accessible error function for all panels.
    """

    def show_error(self):
        sg.PopupError(f'''- Type: {sys.exc_info()[0]}\n- Details: {sys.exc_info()[1]}''')
        logging.error("Exception occurred", exc_info=True)

def main():

    """[summary]
    First page, choose file
    """

    sg.theme('Reddit')

    with open('history.log', 'r') as h:
        lastdata = h.readlines()[-1].split(', ')[1]

    layout = [[sg.Text('File '), sg.InputText(f'{lastdata}', key='_FILEPATH_', s=40), sg.FileBrowse('Browse', target = '_FILEPATH_')],\
        [sg.Text('Skip'), sg.InputText('0', s=3, key='_ROWSKIP_')],
        [sg.OK(s=10)]]

    window = sg.Window('Dataform entry', layout)

    while True:
        event, values = window.read() 
        
        if event == 'OK':
            try:
                homepage(values['_FILEPATH_'], int(values['_ROWSKIP_']))
                break
            except:
                error_func.show_error()
        if event is None:
            break
    
    window.close()

def homepage(_path, rowskip):
    
    '''
    Main page for data entry
    '''
    
    _folder = '/'.join(_path.split('/')[:-1])
    _extension = _path.split('.')[-1]

    df = pd.read_excel(_path, dtype={'hn':'int'}, skiprows=rowskip)

    print('Form launched...')
    print(f'Read data from {_path}')

    all_col = [x for x in df.columns]
    numeric_col = [x for x in df.select_dtypes(['int', 'float']).columns]
    object_col = [x for x in df.select_dtypes('object').columns]
    date_col = [x for x in df.select_dtypes('datetime').columns]

    #? Change to date-only format, dont known why cannot use apply to all columns
    for col in date_col:
        df[f'{col}'] = df[f'{col}'].apply(lambda x: x.date())

    #* ==== Layout =====

    layout1 = [[sg.Text('Numeric column', font=('Arial', 10, 'bold'))],
    [sg.Text('index', size = (5,1)), sg.InputText(size = (4,1), key = 'index'),\
        sg.Button('Get index', pad = (47,0))],
    [sg.Text('hn', size = (3,1)), sg.InputText(size = (8,1), key = 'hn', pad = (20,0)), sg.Button('Get HN')]]

    for variable in numeric_col:
        if (variable == 'index') | (variable == 'hn'):
            continue
        else:
            layout1.append([sg.Text(f'{variable}', size = (8,1)), \
                sg.InputText(size = (20,1), key = f'{variable}',)])

    layout2 = [[sg.Text('Character column', font=('Arial', 10, 'bold'))]]
    for variable in object_col:
        layout2.append([sg.Text(f'{variable}', size = (8,1)), \
            sg.InputText(size = (20,1), key = f'{variable}')])

    layout3 = [[sg.Text('Date column', font=('Arial', 10, 'bold'))]]
    for variable in date_col:
        layout3.append([sg.Text(f'{variable}', size = (8,1)), \
            sg.InputText(size = (20,1), key = f'{variable}')])

    whole_column = [[sg.Column(layout1,vertical_alignment='top'), \
        sg.Column(layout2,vertical_alignment='top'),\
             sg.Column(layout3,vertical_alignment='top')]]

    layout = [[sg.Column(whole_column, size = (800,520), scrollable=True, vertical_scroll_only=True)],
    [sg.Stretch(),\
    sg.Button(image_filename = './pic/ButtonGraphics/First.png', key='_First_', tooltip='First record'), \
    sg.Button(image_filename = './pic/ButtonGraphics/Previous.png', key = '_Previous_', tooltip='Previous record'), \
    sg.Button(image_filename = './pic/ButtonGraphics/New.png', key = '_New_',tooltip='New record'),\
    sg.Button(image_filename = './pic/ButtonGraphics/Next2.png', key='_Next_', tooltip='Next record'),\
    sg.Button(image_filename = './pic/ButtonGraphics/Last.png', key='_Last_', tooltip='Last record'),\
    sg.Stretch()], 
    [sg.Button('Update', tooltip = 'Update data in the form', button_color='green'), \
    sg.Button('Delete', tooltip = 'DELETE THIS INDEX',button_color='red'),\
    sg.Button('Clear', tooltip = 'Clear output'),\
    sg.Stretch(),\
    sg.FileSaveAs(button_text='Save as', file_types=(('xlsx','.xlsx'),('csv','.csv')), default_extension='.xlsx',initial_folder = _folder,target='_FILESAVE_'),\
        sg.Input(key = '_FILESAVE_', visible=False,enable_events=True),\
            sg.Button('Show DF', tooltip = 'Show source dataframe'),\
                 sg.Button('History', tooltip = 'History of last records before close')]]
    
    #* ==== Function =====

    def clear_input():
        for key in values:
            if key in ['index','hn','Save as']:
                continue
            else: 
                window[key]('')
        return None

    def get_variable(df):
        window['index'](df.index[0])
        for variable in numeric_col + object_col + date_col:
            window[f'{variable}'](df[f'{variable}'].iloc[0])

    def save_file(filename, df):
        extension = filename.split('.')[-1]
        if extension == 'xlsx':
            df.to_excel(filename, index=False)
        if extension == 'csv':
            df.to_csv(filename, index=False)

    def update_value(df):        
        try:
            new_df = pd.DataFrame(values, index = [0])
            if 'hn' not in df.columns:
                new_df.drop('hn', axis=1, inplace=True)
            new_df = new_df.apply(lambda x: x.str.strip()).replace(['NaN', 'nan', 'None', ''], np.NaN)
            try:
                new_df[numeric_col] = new_df[numeric_col].astype('float')
                if 'hn' in new_df.columns:
                    new_df['hn'] = new_df['hn'].astype('int')
                for col in date_col:
                    new_df[f'{col}'] = new_df[f'{col}'].apply(lambda x: pd.to_datetime(x).date())
            except:
                error_func.show_error()
                return(df)

            old_value = df.iloc[int(values['index']):int(values['index'])+1].to_dict()
            new_value = {key: values[key] for key in all_col}
            
            answer = sg.popup_ok_cancel(f'''

            Confirm data? 

            Old value: 

            {old_value} 

            -----

            Updated value:

            {new_value}

            ''')

            if answer == 'OK':
                if len(df.iloc[int(values['index']):int(values['index'])+1]) == 0:
                    df = df.append(new_df[all_col], ignore_index=True)
                else:
                    df.iloc[[int(values['index'])]] = new_df[all_col].values[0]
                save_file(f'backup/temp.{_extension}', df) 
                return(df)
            else:
                return(df)
        except:
           error_func.show_error()
           return(df)
    
    def delete_row(df):
        answer = sg.popup_ok_cancel('Delete this index?')
        if answer == 'OK':
            answer2 = sg.popup_ok_cancel('Sure?')
            if answer2 == 'OK':
                save_file(f'backup/temp_before_delete.{_extension}', df)
                df = df.drop(int(values['index'])).reset_index(drop=True)

                #? Re-evaluate form
                if df.shape[0] == 0:
                    clear_input()
                elif int(values['index'])+1 > df.shape[0]:
                    get_variable(df.iloc[[0]])
                else:
                    get_variable(df.iloc[[int(values['index'])]])

                save_file(f'backup/temp.{_extension}', df)
            return(df)
        else:
            return(df)

    #* ==== Start application here =====

    window = sg.Window(f'Data entry form {_path}', layout, element_justification='l', grab_anywhere=True,\
         enable_close_attempted_event=True, alpha_channel=0.97,finalize=True)

    while True:             

        event, values = window.read()

        logging.info(f'event: {event}\nvalues: {values}')
             
        if event == 'Get HN':
            try:        
                df_tmp = df.loc[df['hn'] == int(values['hn'])]
                get_variable(df_tmp)
            except:
                error_func.show_error()

        if event == 'Get index':     
            df_tmp = df.iloc[[int(values['index'])]]
            get_variable(df_tmp)

        if event == 'History':
                open_history()
    
        if event == '_First_':
            df_tmp = df.iloc[[0]]
            get_variable(df_tmp)

        if event == '_Last_':
            df_tmp = df.iloc[[-1]]
            get_variable(df_tmp)

        if event == '_Next_':
            try:
                if (values['index']) == '':
                    df_tmp = df.iloc[[0]]
                elif (int(values['index'])+2) > df.shape[0]:
                    df_tmp = df.iloc[[0]]
                else: 
                    df_tmp = df.iloc[[(int(values['index'])+1)]]
                get_variable(df_tmp)    
            except:
                error_func.show_error()

        if event == '_Previous_':
            try:
                if (values['index']) == '':
                    df_tmp = df.iloc[[-1]]
                else:
                    df_tmp = df.iloc[[(int(values['index'])-1)]]
                get_variable(df_tmp)    
            except:
                error_func.show_error()
        
        if event == '_New_':
            window['index'](len(df))
            window['hn']('')
            clear_input()

        if event == 'Delete':
            df = delete_row(df)

        if event == 'Update':
            try:
                df = update_value(df)
            except:
                error_func.show_error()

        if event == 'Clear':
            clear_input()
            
        if event == 'Show DF':
            open_df(df)

        if event == '_FILESAVE_':
            save_file(values['_FILESAVE_'], df)
       
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT:
            choice = before_close(values, _path, _folder)
            if choice[0] == 'exit':
                if choice[1] is None:
                    break
                else:
                    save_file(choice[1], df)
                    break
            if choice[0] == 'continue':
                continue

    print('Close the program...')

    window.close()

def open_df(df):
    
    """
    This function create a dataframe from main df output
    """
    
    header_list = [x for x in df.columns] 

    layout = [[sg.Table(df.values.tolist(), headings=header_list, auto_size_columns=False,\
        vertical_scroll_only=False, num_rows=20, display_row_numbers=True)]]

    window = sg.Window('Dataframe', layout, modal=True, size=(800,400))

    window.read(close=True)
    
def open_history():
    
    """
    This function prints histories of last file + index opened before closed
    """

    text = ''
    with open('history.log', 'r+') as h:
        for line in h:
            text += line

    layout = [[sg.Multiline(text, size = (80,10), key = '_output_')],
    [sg.Button('Clear'), sg.Button('Exit')]]
    
    window = sg.Window('History', layout, modal = True)

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        
        if event == 'Clear':
            answer = sg.PopupOKCancel('Clear history?')
            if answer == 'OK':
                h = open('history.log', 'r+')
                h.truncate(0)
                h.close()
                text = ''
                break
            
    window.close()

def before_close(main_values, main_path, folder):

    """
    This function create pop-up with save as output
    main_values = values from main()
    """

    layout = [[sg.Text('Save before exit?')],\
        [sg.FileSaveAs(button_text='Yes', file_types=(('xlsx','.xlsx'),('csv','.csv')), \
        default_extension='.xlsx',initial_folder = folder, target='_SAVEEXIT_', s=5),\
            sg.Input(key = '_SAVEEXIT_', visible=False, enable_events=True),\
                 sg.No(s=5), sg.Cancel(s=5)]]
    
    event, values = sg.Window('Exit confirmation', layout, modal=True, \
        disable_minimize=True, disable_close=True).read(close=True)

    if event == '_SAVEEXIT_':
        if main_values['index'] != '':
            with open('history.log', 'a') as h:
                h.write(f'{date.today()}, {main_path}, {main_values["index"]}\n')
        return(['exit', values['_SAVEEXIT_']])

    if event == 'No':
        if main_values['index'] != '':
            with open('history.log', 'a') as h:
                h.write(f'{date.today()}, {main_path}, {main_values["index"]}\n')
        return(['exit', None])
    if event == 'Cancel':
        return(['continue'])
    if event is None:
        return(['continue'])

if __name__ == '__main__':
    path = Path(__file__).parent.absolute()
    os.chdir(path)

    logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',\
    datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO, \
        handlers=[logging.FileHandler(filename='app.log', mode='a',  encoding='utf-8')])

    error_func = error_handler()
    main()
