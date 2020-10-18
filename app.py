import PySimpleGUI as sg
import sqlite3
import subprocess
import datetime

def scrapePopup(site):
    layout = [
        [
            sg.Text('DB file'), sg.InputText(), sg.FileBrowse()
        ],
        [
            sg.Text('Roots file'), sg.InputText(), sg.FileBrowse()
        ],
        [
            sg.Text('Proxies file'), sg.InputText(), sg.FileBrowse()
        ],
        [
            sg.Button('Load'), sg.Cancel()
        ]
    ]

    window = sg.Window(title='', layout=layout)

    while(True):
        event, values = window.read()

        if event in (None, 'Exit', 'Cancel'):
            break
        if event == 'Load':
            if(site == 1):
                subprocess.Popen(['python', 'scraper/optionisticScraper.py', '-db' , values[0], '-roots', values[1], '-proxies', values[2]])
            break
        
    window.close()


def baseWindow():
    layout = [
        [
            sg.Button('Scrape Optionistics')
        ]
    ]
    return sg.Window(title='Nasdaq scraper', layout=layout, resizable=True)



def main():
    db = None
    tables = None
    table = None
    lag = []

    sg.theme('SystemDefault')
    mainWindow = baseWindow()

    while(True):
        event, values = mainWindow.read()

        if event in (None, 'Exit', sg.WIN_CLOSED):
            break
        if event == 'Load DB':
            if(values[0] != ''):
                try:
                    db = sqlite3.connect(values[0])
                    cursor = db.cursor()
                except:
                    print('Database connection failed')
                else:
                    print('Database connected succesfully')
                    try:
                        tables = getTables(cursor)
                    except sqlite3.DatabaseError as err:       
                        print("Error: ", err)
                    else:
                        buffer = windowWithTables(tables)
                        mainWindow.close()
                        mainWindow = buffer

        if event == 'Scrape Nasdaq':
            scrapePopup(0)

        if event == 'Scrape Optionistics':
            scrapePopup(1)

    if(db):
        db.close()
        
    mainWindow.close()
if __name__ == "__main__":
    main()