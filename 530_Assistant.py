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
        row = TX_Table.GetValues([nrow], list(range(1, TX_Table.ColumnCount + 10)))
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


def get_calculated_value_bilin_interp(file):
    ppt = Get_Selected_Link_Properties()
    lon = (ppt["ABS_X_A"] + ppt["ABS_X_B"]) / 2
    lat = (ppt["ABS_Y_A"] + ppt["ABS_Y_B"]) / 2
    lon_lat, values = find_4_closest_MW_Calculated_Value(
        file, lat, lon)
    return bilinear_interpolation(lat, lon, lon_lat, values)


def AtollMacro_K(printed=True):
    ppt = Get_Selected_Link_Properties()
    lon = (ppt["ABS_X_A"] + ppt["ABS_X_B"]) / 2
    lat = (ppt["ABS_Y_A"] + ppt["ABS_Y_B"]) / 2
    res = get_calculated_value_bilin_interp(os.path.join(path, "530-18_calculated_values\\LogK_merged.csv"))
    if printed:
        print("K = " + str(10**res))
    return 10**res


def AtollMacro_p0():
    # eps_p
    ppt = Get_Selected_Link_Properties()
    print("Don't forget to fill Altitudes.txt file from Profile Values")
    file = open(os.path.join(path, "530-18_calculated_values\\Altitudes.txt"), 'r')
    content = file.readlines()
    file.close()
    dN75 = get_calculated_value_bilin_interp(os.path.join(path, "530-18_calculated_values\\dN75_merged.csv"))
    h_t = 0
    d = ppt["LINK_LENGTH"] / 1000
    h_e = float(content[0]) + ppt["HEIGHT_A"]
    h_r = float(content[-1]) + ppt["HEIGHT_B"]
    h_L = min(h_e, h_r)
    # (5)
    eps_p = np.abs(h_r - h_e) / d
    f = ppt["FREQ_A"] / 1000
    for r in content:
        h_t += float(r.replace("\n", ""))
    h_t /= len(content)
    # (6)
    h_c = (h_e + h_r) / 2 - d**2 / 102 - h_t
    # (8), (9)
    v_sr = min((dN75 / 50)**1.8 * np.exp(-(h_c / (2.5 * np.sqrt(d)))), dN75 * d**1.5 * f**0.5 / 24730)
    # (11)
    p0 = AtollMacro_K(False) * d**3.51 * (f**2 + 13)**0.447 * 10**(-0.376 * np.tanh((h_c - 147) / 125) - 0.334 * eps_p **
                                                                   0.39 - 0.00027 * h_L + 17.85 * v_sr)
    print(f"p0 : {p0}%")
    return p0


def AtollMacro_Print_Selected_Link_Properties():
    print = print_to_window
    print(Get_Selected_Link_Properties())


lon = -61.593341667
lat = 16.225
