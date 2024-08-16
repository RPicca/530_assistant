import win32com.client
import PySimpleGUI as sg
import os
import sys

path = "D:\Documents\Atoll_Macros"


def print_to_window(message):
    if isinstance(message, dict):
        for k in message.keys():
            message[k] = str(message[k])
        layout = [
            [sg.Table(
                headings=list(message.keys()),
                values=[list(message.values())],
                vertical_scroll_only=False)]]
    else:
        layout = [[sg.Text(str(message))]]
    # Create the window
    window = sg.Window('Printing Window', layout, grab_anywhere=False, resizable=True)
    while True:
        # Display and interact with the Window
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
    window.close()


def Custom_report(dir_path):
    doc = Atoll.ActiveDocument
    links = doc.GetRootFolder(0).Item("Links").Item("Microwave Links")
    inputs = doc.GetCommandDefaults("MWLinksCustomExport")
    inputs.Set("EXPORTPATH", dir_path)
    inputs.Set(
        "TEMPLATEFILEPATH",
        "D:\Documents\Temp\CR_530-18.txt")
    inputs.Set("LINKSFOLDER", links)
    output = doc.InvokeCommand("MWLinksCustomExport", inputs)


def Get_Selected_Link_Properties():
    doc = Atoll.ActiveDocument
    link = doc.Selection
    if link == None:
        return None
    else:
        TX_Table = win32com.client.dynamic.Dispatch(doc.GetRecords("MWLinks", True))
        nrow = TX_Table.FindPrimaryKey(link.Name)
        row = TX_Table.GetValues([nrow], list(range(1, TX_Table.ColumnCount + 1)))
        dico = {}
        for i in range(1, len(row[0]) - 3):
            dico[row[0][i]] = row[1][i]
        return dico

# TODO : utilisation mixe du CR et des infos de la table des liaisons MW pour écrire tous les fichiers à lire par Smath


def AtollMacro_write_smath_files():
    dir_path = os.path.join(path, "smath_files")
    dico = Get_Selected_Link_Properties()
    link_name = Atoll.ActiveDocument.Selection.Name
    if not os.path.isdir(dir_path):
        os.mkdir(dir_path)
    for k in dico.keys():
        with open(os.path.join(dir_path, k + ".txt"), "w") as file:
            file.write(str(dico[k]))
    Custom_report(dir_path)
    print(read_CR(link_name))


def read_CR(link_name):
    dico = {}
    with open(link_name + ".txt", "r") as file:
        for r in file.readlines:
            dico[r.split[": "][0]] = r.split[": "][1]
    return dico


def AtollMacro_Print_Selected_Link_Properties():
    print = print_to_window
    print(Get_Selected_Link_Properties())
