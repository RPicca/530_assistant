import win32com.client
import PySimpleGUI as sg
import os
import openpyxl
import numpy as np

path = "\\\sisyphe\\TestsStorage\\Tests_Atoll\\Traitements\\MW\\Feuilles de calculs SMath\\530 - Interruptions dues à la pluie et aux multi-trajets"


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
    dir_path = os.path.join(dir_path, "CR")
    if not os.path.isdir(dir_path):
        os.mkdir(dir_path)
    doc = Atoll.ActiveDocument
    links = doc.GetRootFolder(0).Item("Links").Item("Microwave Links")
    inputs = doc.GetCommandDefaults("MWLinksCustomExport")
    inputs.Set("EXPORTPATH", dir_path)
    inputs.Set(
        "TEMPLATEFILEPATH",
        os.path.join(dir_path, "Template_CR_530-18.txt"))
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
    dico_CR = read_CR(os.path.join(dir_path, "CR", link_name))
    for k in dico_CR:
        with open(os.path.join(dir_path, k + ".txt"), "w") as file:
            # Useful for min phase radio params
            if dico_CR[k].split(" ")[0] == "1e-6":
                string = dico_CR[k].split(" ")[1]
            # Remove digit separator and Unit and other stuff...
            else:
                string = str(dico_CR[k]).split(" ")[0].replace(",", "").replace(
                    "Yes", "1").replace("No", "0").replace("<Ignore>", "").replace("n/a", "0")
            try:
                float(string)
            except:
                if string == "":
                    string = "0"
                else:
                    string = "\"" + string + "\""
            file.write(string)


def read_CR(link_name):
    dico = {}
    with open(link_name + ".txt", "r") as file:
        for r in file.readlines():
            dico[r.split(": ")[0]] = r.split(": ")[1].replace("\n", "")
    return dico


def find_4_closest_MW_Calculated_Value(file, lat, lon):
    table = open(file, 'r')
    content = []
    for l in table.readlines():
        list = [float(i) for i in l.split(";")]
        content.append(list)
    table.close()
    i = j = 1
    while (content[0][j] - lon) * (content[0][j + 1] - lon) > 0:
        j += 1
    while (content[i][0] - lat) * (content[i + 1][0] - lat) > 0:
        i += 1
    lon_lat = [[content[0][j], content[0][j + 1]],
               [content[i][0], content[i + 1][0]]]
    values = [[content[i][j], content[i][j + 1]],
              [content[i + 1][j], content[i + 1][j + 1]]]
    return lon_lat, values


def linear_interpolation(x, coord, values):
    if coord[0] > coord[1]:
        coord.reverse()
        values.reverse()
    slope = (values[1] - values[0]) / (coord[1] - coord[0])
    return values[0] + slope * np.abs(x - coord[0])


def bilinear_interpolation(lat, lon, lon_lat, values):
    tmp = [linear_interpolation(lon, lon_lat[0], values[0]), linear_interpolation(lon, lon_lat[0], values[1])]
    return linear_interpolation(lat, lon_lat[1], tmp)


def AtollMacro_K():
    ppt = Get_Selected_Link_Properties()
    lon = (ppt["ABS_X_A"] + ppt["ABS_X_B"]) / 2
    lat = (ppt["ABS_Y_A"] + ppt["ABS_Y_B"]) / 2
    lon_lat, values = find_4_closest_MW_Calculated_Value(
        "\\\sisyphe\\TestsStorage\\Tests_Atoll\\Traitements\MW\\Feuilles de calculs SMath\\530 - Interruptions dues à la pluie et aux multi-trajets\\530-18_calculated_values\\LogK_merged.csv",
        lat, lon)
    print("K = " + str(10**bilinear_interpolation(lat, lon, lon_lat, values)))


def AtollMacro_Print_Selected_Link_Properties():
    print = print_to_window
    print(Get_Selected_Link_Properties())


lon = -6.1806828726247485
lat = 53.288487911977015
lon_lat, values = find_4_closest_MW_Calculated_Value(
    "\\\sisyphe\\TestsStorage\\Tests_Atoll\\Traitements\MW\\Feuilles de calculs SMath\\530 - Interruptions dues à la pluie et aux multi-trajets\\530-18_calculated_values\\LogK_merged.csv",
    lat, lon)

K = 10**bilinear_interpolation(lat, lon, lon_lat, values)

print(K)
