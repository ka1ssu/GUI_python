from openpyxl import load_workbook
from datetime import datetime
import PySimpleGUI as sg

layout = [[sg.Text('ФИО'), sg.Push(), sg.Input(key='FIO')],
          [sg.Text('Номер телефона'), sg.Push(), sg.Input(key='telephone')],
          [sg.Text('Название'), sg.Push(), sg.Input(key='name')],
          [sg.Text('Производитель'), sg.Push(), sg.Input(key='prod')],
          [sg.Text('Количество'), sg.Push(), sg.Input(key='kolvo')],
          [sg.Button('Добавить'), sg.Button('Закрыть')]]
window = sg.Window('info', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Закрыть":
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('info.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            data = [ID, values['FIO'], values['telephone'], values['name'], values['prod'], values['kolvo'], time_stamp]
            sheet.append(data)
            wb.save('info.xlsx')
            window['FIO'].update(value='')
            window['telephone'].update(value='')
            window['name'].update(value='')
            window['prod'].update(value='')
            window['kolvo'].update(value='')
            window['name'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')
window.close()
