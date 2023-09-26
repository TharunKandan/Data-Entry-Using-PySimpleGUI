from pathlib import Path
import PySimpleGUI as sg
import pandas as pd


sg.theme('Darkteal9')

EXCEL_FILE=Path('Data_Entry.xlsx')
if EXCEL_FILE.exists():
    df=pd.read_excel(EXCEL_FILE)
else:
    df=pd.DataFrame()

layout = [
    [sg.Text('Please fill the below fields correctly')],
    [sg.Text('Name:', size=(15, 1)), sg.InputText(key='Name')],
    [sg.Text('Reg no:', size=(15,1)), sg.InputText(key='Reg no')],
    [sg.Text('Vm no:', size=(15,1)), sg.InputText(key='Vm no')],
    [sg.Text('Department:',size=(15,1)), sg.Combo(['B.TEHC','B.E'], key='Dept')],
    [sg.Text('Branch:',size=(15,1)), sg.Combo(['Artificial Intelligence and Data Science','Mechanical Engineering','Civil Engineering','Computer Science and Engineering','Computer Science and Business systems','Information Technology'], key='Branch')],
    [sg.Text('Year:', size=(15,1)), sg.Spin([i for i in range(1,5)],
                                           initial_value=1, key='Year')],
    [sg.Text('Semester:',size=(15,1)), sg.Combo(['I','II','III','IV','V','VI','VII','VIII'], key='Semester')],
    [sg.Text('Address:', size=(15, )), sg.InputText(key='Address')],
    [sg.Text('Phone Number', size=(15, 1)), sg.InputText(key='Phno')],
    [sg.Text('Email Id', size=(15, 1)), sg.InputText(key='EmId')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

# Create the window
window = sg.Window('Student Data Entry', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df=pd.concat([df,new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()

window.close()
