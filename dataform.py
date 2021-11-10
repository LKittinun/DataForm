import PySimpleGUI as sg
import pandas as pd
import numpy as np
import os
import sys
import logging
from datetime import date
from pathlib import Path

path = Path(__file__).parent.absolute()
os.chdir(path)

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',\
    datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO, \
        handlers=[logging.FileHandler(filename='app.log', mode='a',  encoding='utf-8')])

def main():

    sg.theme('DarkBlue 14')

    _path = r'C:/Users/Kittinun/Desktop/R Workspace/CCRT/test.xlsx'
    _folder = '/'.join(_path.split('/')[:-1])
    
    df = pd.read_excel(_path)

    print('Form launched...')
    print(f'Read data from {_path}')

    all_col = [x for x in df.columns]
    numeric_col = [x for x in df.select_dtypes(['int', 'float']).columns]
    object_col = [x for x in df.select_dtypes('object').columns]
    date_col = [x for x in df.select_dtypes('datetime').columns]

    #* Change to date-only format, dont known why cannot use apply to all columns
    for col in date_col:
        df[f'{col}'] = df[f'{col}'].apply( lambda x: x.date())

    layout1 = [[sg.Text('Numeric column', font=('Arial', 10, 'bold'))],
    [sg.Text('index', size = (5,1)), sg.InputText('-1', size = (4,1), key = 'index'),\
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
    [sg.Stretch(),sg.Button(image_filename = './pic/ButtonGraphics/First.png', key='_First_'), \
        sg.Button(image_filename = './pic/ButtonGraphics/Previous.png', key = '_Previous_'), \
            sg.Button(image_filename = './pic/ButtonGraphics/Next2.png', key='_Next_'), \
                sg.Button(image_filename = './pic/ButtonGraphics/Last.png', key='_Last_'),\
                    sg.Stretch()], 
    [sg.Button('Update'), sg.Button('Clear'), sg.Stretch(),
    sg.FileSaveAs('Save file', enable_events=True, file_types=(('xlsx', '*.xlsx'), ('csv', '.csv')),\
        default_extension='.xlsx', key = '_FILE_', initial_folder=_folder), \
            sg.Button('Show DF'), sg.Button('History')]]


    def clear_input():
        for key in values:
            if key in ['index','hn','_FILE_']:
                continue
            else: 
                window[key]('')
        return None

    def get_variable(df):
        window['index'](df.index[0])
        for variable in numeric_col + object_col + date_col:
            window[f'{variable}'](df[f'{variable}'].iloc[0])

    def show_error():
        sg.PopupError(f'''- Type: {sys.exc_info()[0]}\n- Details: {sys.exc_info()[1]}''')
        logging.error("Exception occurred", exc_info=True)

    def save_file():
        filename = values['_FILE_']
        print(filename)
        extension = filename.split('.')[-1]
        print(extension)
        if extension == 'xlsx':
            df.to_excel(filename, index=False)
        if extension == 'csv':
            df.to_csv(filename, index=False)

    def update_value(df):        
        try:
            new_df = pd.DataFrame(values, index = [0])
            new_df = new_df.apply(lambda x: x.str.strip()).replace(['NaN', 'nan', 'None', ''], np.NaN)
            try:
                new_df[numeric_col] = new_df[numeric_col].astype('float')
                for col in date_col:
                    new_df[f'{col}'] = new_df[f'{col}'].apply(lambda x: pd.to_datetime(x).date())
            except:
                show_error()
                return
            old_value = df.loc[df['hn'] == int(values['hn'])].to_dict()
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
                df.loc[df['hn'] == int(values['hn'])] = new_df[all_col].values[0]
        except :
           show_error()

    window = sg.Window(f'Data entry form {_path}', layout, element_justification='l', grab_anywhere=True,\
         enable_close_attempted_event=True)



    while True:             

        event, values = window.read()

        logging.info(f'event: {event}\nvalues: {values}')

        if event == 'Get HN':
            try:        
                df_tmp = df.loc[df['hn'] == int(values['hn'])]
                get_variable(df_tmp)
            except:
                show_error()

        if event == 'Get index':
            try:        
                df_tmp = df.iloc[[int(values['index'])]]
                get_variable(df_tmp)
            except:
                show_error()

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
                if (int(values['index'])+2) > df.shape[0]:
                    df_tmp = df.iloc[[0]]
                else: 
                    df_tmp = df.iloc[[(int(values['index'])+1)]]
                get_variable(df_tmp)    
            except:
                show_error()

        if event == '_Previous_':
            try:
                df_tmp = df.iloc[[(int(values['index'])-1)]]
                get_variable(df_tmp)    
            except:
                show_error()

        if event == 'Update':
            update_value(df)

        if event == 'Clear':
            clear_input()
            
        if event == 'Show DF':
            open_df(df)

        if event == '_FILE_':
            save_file()
       
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT:
            answer = sg.popup_ok_cancel('Are you sure?\nPlease save the file.')
            if answer == 'OK':
                if values['index'] != '-1':
                    with open('history.log', 'a') as h:
                        h.write(f'{date.today()}: {values["index"]}\n')
                break
            else:
                continue
    
    print('Close the program...')

    window.close()

def open_df(df):
    
    header_list = [x for x in df.columns] 

    layout = [[sg.Table(df.values.tolist(), headings=header_list, auto_size_columns=False,\
        vertical_scroll_only = False, num_rows=20,display_row_numbers=True)]]

    window = sg.Window('Dataframe', layout, modal=True, size=(800,500))

    window.read()
    
    window.close()

def open_history():
    
    text = ''
    with open('history.log', 'r+') as h:
        for line in h:
            text += line

    layout = [[sg.Multiline(text, size = (50,10), key = '_output_')],
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

if __name__ == '__main__':
    main()
