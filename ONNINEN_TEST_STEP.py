import json
import os
import urllib.request
import warnings
from datetime import datetime
from enum import Enum
from pathlib import Path
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askdirectory
from urllib.error import HTTPError
from unidecode import unidecode
import pandas
import requests
from requests.auth import HTTPBasicAuth
from ttkthemes.themed_style import ThemedStyle

warnings.filterwarnings('ignore')
df = pandas.DataFrame()
final_df = pandas.DataFrame()
i = 0
now = datetime.now()
todaydate = now.strftime("%d-%m-%Y--%H.%M.%S")
bledy_zatwierdzenia = []
filepath_local = ''
wersja_prog = '0.10'

# FIXME Wgrano wyświetlanie błędu autoryzacji której nie było w wersji 0.7
#      Wersja 0.9 Poprawiono nazwy assetow, brak spacji przed/za konwersja polskich znakow
#      Wersja 0.10 dodano kolor dla błędnych assetów.




k = 0


class AppURLopener(urllib.request.FancyURLopener):
    version = "Mozilla/5.0"

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

class ProgramStatus(Enum):
    ProgramStart = 0
    AssetCreated = 1
    AssetUpdated = 2
    AssetApproved = 3
    DataFrame = 4
    AssetLinked = 5
    ProductApproved = 6


def createAsset(step_folder, assetid, obj_type, date, document_type, certyficate_type, username, password):
    if pandas.isna(date):
        date = ""

    if pandas.isna(certyficate_type):
        certyficate_type = ""

    url = f"https://steppimprod001.ku.k-netti.com/restapiv2/assets/{assetid}?allow-overwrite=false&context=pl-PL&workspace=Main"
    headers = {'Content-Type': 'application/json', 'accept': 'application/json'}
    dane = {
        "name": str(assetid),
        "objectType": obj_type,
        "classifications": [
            str(step_folder)
        ],
        "values": {
            "Asset_Validity_Date": {
                "value": {
                    "value": date
                }
            },
            "PL_asset_push": {
                "value": {
                    "value": "YES"
                }
            },
            "PL_DocumentTypeOther": {
                "value": {
                    "value": document_type
                }
            },
            "PL_DocumentType": {
                "value": {
                    "value": certyficate_type
                }
            }
        }
    }
    return requests.put(url,
                        auth=HTTPBasicAuth(username, password),
                        data=json.dumps(dane),
                        headers=headers,
                        verify=False)


def updateAsset(assetid, files, username, password):
    url_update = f"https://steppimprod001.ku.k-netti.com/restapiv2/assets/{assetid}/content?fileName=unknown&context=pl-PL&workspace=Main"
    headers = {'Content-Type': 'application/octet-stream', 'accept': '*/*'}
    return requests.put(url_update,
                        auth=HTTPBasicAuth(username, password),
                        data=files,
                        headers=headers,
                        verify=False)


def approveAsset(assetid, username, password):
    headers = {'accept': 'application/json'}
    url = f"https://steppimprod001.ku.k-netti.com/restapiv2/assets/{assetid}/approve?context=pl-PL&workspace=Main"
    return requests.post(url,
                         auth=HTTPBasicAuth(username, password),
                         headers=headers,
                         verify=False)


def linkAsset(pimid, username, password, id_reference, assetid):
    data = '{}'
    headers = {'Content-Type': 'application/json', 'accept': 'application/json'}
    url = f"https://steppimprod001.ku.k-netti.com/restapiv2/products/{pimid}/references/{id_reference}/{assetid}?allow-overwrite=true&context=pl-PL&workspace=Main"
    return requests.put(url,
                        auth=HTTPBasicAuth(username, password),
                        data=data,
                        headers=headers,
                        verify=False)


def approveProduct(pimid, username, password):
    url = f"https://steppimprod001.ku.k-netti.com/restapiv2/products/{pimid}/approve?context=pl-PL&workspace=Main"
    return requests.post(url,
                         auth=HTTPBasicAuth(username, password),
                         verify=False)


def get_file():
    global df
    filepath = askopenfilename()
    df = pandas.read_excel(filepath)
    df['NAZWA'] = df['NAZWA'].apply(str)


def get_file_local():
    global filepath_local
    filepath_local = askdirectory()


def start():
    submitForm(filepath_local)


def program_link(step_folder, assetid, obj_type, date, document_type, certyficate_type, username, url_asset,
                 id_reference, pimid, password, k, c):
    retries = 5
    programStage = ProgramStatus.ProgramStart
    response_url = requests.get(url_asset, headers={'User-Agent': 'Mozilla/5.0'}, stream=True)
    files = response_url.content
    df['LOG'].astype(str)

    for trying in range(retries):
        if programStage == ProgramStatus.ProgramStart:
            response = createAsset(step_folder, assetid, obj_type, date, document_type, certyficate_type, username,
                                   password)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetCreated
            elif response.status_code == 400:
                programStage = ProgramStatus.AssetApproved
            elif response.status_code == 401:
                print(f"{bcolors.FAIL} Błędny login lub hasło {bcolors.ENDC}")
                obj_type = '!!!Bledny login lub hasło!!!!'
                print("Blad tworzenia assetu")
                progress_bar(obj_type, c, df)
            else:
                print("Blad tworzenia assetu")
                df.at[k, 'LOG'] = "Asset nie stworzony!"
                continue
        if programStage == ProgramStatus.AssetCreated:
            response = updateAsset(assetid, files, username, password)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetUpdated
            else:
                print(response.status_code)
                print(response.content)
                df.at[k, 'LOG'] = "Asset stworzony, zawartość nie wgrana!"
                continue
        if programStage == ProgramStatus.AssetUpdated:
            response = approveAsset(assetid, username, password)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetApproved
            else:
                df.at[k, 'LOG'] = "Asset bez approve!"
                continue
        if programStage == ProgramStatus.AssetApproved:
            response = linkAsset(pimid, username, password, id_reference, assetid)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetLinked
            elif response.status_code == 401:
                print(f"{bcolors.FAIL} Błędny login lub hasło {bcolors.ENDC}")
                obj_type = '!!!Bledny login lub hasło!!!!'
                progress_bar(obj_type, c, df)
            else:
                df.at[k, 'LOG'] = "Asset nie przypisany do indeksu!"
                continue
        if programStage == ProgramStatus.AssetLinked:
            response = approveProduct(pimid, username, password)
            if 200 <= response.status_code < 300:
                print(f"{assetid} Wgrany na indeksie {pimid} {response.status_code}")
                df.at[k, 'LOG'] = "Wgrano"
                break


def program_folder(step_folder, assetid, obj_type, date, document_type, certyficate_type, username, files, id_reference,
                   pimid, password, k, new_df, c, df):
    retries = 5
    programStage = ProgramStatus.ProgramStart
    new_df['LOG'] = new_df['LOG'].astype("string")
    for trying in range(retries):
        if programStage == ProgramStatus.ProgramStart:
            response = createAsset(step_folder, assetid, obj_type, date, document_type, certyficate_type, username,
                                   password)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetCreated
            elif response.status_code == 400:
                print("Asset istnieje")
                programStage = ProgramStatus.AssetApproved
            elif response.status_code == 401:
                print(f"{bcolors.FAIL} Błędny login lub hasło {bcolors.ENDC}")
                obj_type = '!!!Bledny login lub hasło!!!!'
                progress_bar(obj_type, c, df)
            else:
                print(response.status_code, response.content)
                continue
        if programStage == ProgramStatus.AssetCreated:
            response = updateAsset(assetid, files, username, password)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetUpdated
            else:
                print(response.status_code, response.content)
                continue
        if programStage == ProgramStatus.AssetUpdated:
            response = approveAsset(assetid, username, password)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetApproved
            else:
                print(response.status_code, response.content)
                continue
        if programStage == ProgramStatus.AssetApproved:
            response = linkAsset(pimid, username, password, id_reference, assetid)
            if 200 <= response.status_code < 300:
                print(response.status_code)
                programStage = ProgramStatus.AssetLinked
            elif response.status_code == 401:
                print(f"{bcolors.FAIL} Błędny login lub hasło {bcolors.ENDC}")
                obj_type = '!!!Bledny login lub hasło!!!!'
                progress_bar(obj_type, c, df)
            else:
                print(response.status_code, response.content)
                continue
        if programStage == ProgramStatus.AssetLinked:
            response = approveProduct(pimid, username, password)
            if 200 <= response.status_code < 300:
                new_df['LOG'] = new_df['LOG'].astype("string")
                ok = 'Wgrano'
                new_df.at[k, 'LOG'] = str(ok)
                print(f"{assetid} Wgrany na indeksie {pimid} {response.status_code}")
                break
            else:
                print(response.status_code, response.content)
    else:
        new_df.at[k, 'LOG'] = f"ERROR"


def progress_bar(obj_type, c, df):
    progres = ttk.LabelFrame(app, text="Status", height='85', width='290')
    progres.grid(row=1, column=1, padx=5, pady=0, ipadx=50, ipady=0, sticky='W')
    procesing_line = ttk.Label(progres, text=f"{obj_type} {c + 1} z {len(df)}")
    procesing_line.place(x=35, y=20)
    procesing_line.update()


def submitForm(filepath_local):
    global final_df
    upload = ratio.get()
    if upload == 1:
        c = 0
        k = 0

        for i in range(len(df)):
            print(df)
            print(i)
            print(len(df))
            reference_type = df.at[i, 'TYP_REFERENCJI']
            document_type = df.at[i, 'TYP_DOKUMENTU']
            assetid = df.at[i, 'NAZWA']
            if reference_type == "Zdjęcie Główne":
                assetid = "ZG_" + assetid
            elif reference_type == "Zdjęcie Dodatkowe":
                assetid = "ZD_" + assetid
            elif reference_type == "Zdjęcie Wymiarowania":
                assetid = "ZW_" + assetid
            elif reference_type == "Zdjęcie Techniczne":
                assetid = "ZT_" + assetid
            elif reference_type == "Zdjęcie Główne Polska":
                assetid = "ZG_PL_" + assetid
            elif reference_type == "Dokumenty":
                assetid = "DOC_" + assetid
            elif reference_type == "Certyfikaty":
                assetid = "CERT_" + assetid
            assetid = assetid[0:40]
            assetid = assetid.strip()
            certyficate_type = df.at[i, 'TYP_CERTYFIKATU']
            id_reference = reference_dictionary[reference_type]
            step_folder = df.at[i, 'ID_FOLDERU']
            username = loginEntry.get()
            password = passwordEntry.get()
            obj_type = "JPG Image"
            pimid = df.at[i, 'PIM']
            date = df.at[i, 'DATA YYYY-MM-DD']
            url_asset = df.at[i, 'LINK']
            if id_reference == 'PL_Documents' or id_reference == 'PL_Certificates_Reference':
                obj_type = "PDF"

            # Progress Bar
            progress_bar(obj_type, c, df)

            try:
                url_asset = url_asset.replace(" ", "")
                response = requests.get(url_asset, headers={'User-Agent': 'Mozilla/5.0'}, stream=True)

            except HTTPError:
                df.at[k, 'LOG'] = "Błędny link"
                c += 1
                k += 1
            else:
                program_link(step_folder, assetid, obj_type, date, document_type, certyficate_type, username, url_asset,
                             id_reference, pimid, password, k, c)
                c += 1
                k += 1

            if i == len(df) - 1:
                progress_bar(obj_type, c, df)
                df.to_excel(f"LOG_LINK{todaydate}.xlsx", index=False)

    elif upload == 2:
        username = loginEntry.get()
        password = passwordEntry.get()
        i = 0
        for file in os.scandir(filepath_local):
            k = 0
            path_to_file = os.fsdecode(file)
            filename = Path(file).stem
            print(filename)
            new_df = df[df.eq(filename).any(1)]
            new_df.reset_index(drop=True, inplace=True)
            print(new_df)
            if new_df.empty:
                print(f"{bcolors.WARNING} Plik {filename} nie wystepuje w pliku excel sprawdź poprawność nazwy {bcolors.ENDC}")
                continue
            else:
                for b in range(len(new_df)):
                    assetid = unidecode(new_df.at[b, 'NAZWA'])
                    assetid = assetid[0:40]
                    assetid = assetid.strip()
                    obj_type = "JPG Image"
                    pimid = new_df.at[b, 'PIM']
                    reference_type = new_df.at[b, 'TYP_REFERENCJI']
                    document_type = new_df.at[b, 'TYP_DOKUMENTU']
                    certyficate_type = new_df.at[b, 'TYP_CERTYFIKATU']
                    date = new_df.at[b, 'DATA YYYY-MM-DD']
                    id_reference = reference_dictionary[reference_type]
                    step_folder = new_df.at[b, 'ID_FOLDERU']
                    files = open(path_to_file, 'rb').read()
                    if id_reference == 'PL_Documents' or id_reference == 'PL_Certificates_Reference':
                        obj_type = "PDF"

                    # Progress Bar
                    progres = ttk.LabelFrame(app, text="Status", height='85', width='290')
                    progres.grid(row=1, column=1, padx=5, pady=0, ipadx=50, ipady=0, sticky='W')
                    procesing_line = ttk.Label(progres, text=f"Ladowanie {filename}")
                    procesing_line.place(x=35, y=20)
                    procesing_line.update()
                    c = k
                    program_folder(step_folder, assetid, obj_type, date, document_type, certyficate_type, username,
                                   files,
                                   id_reference, pimid, password, k, new_df, df, c)
                    i += 1
                    k += 1
                final_df = final_df.append(new_df)

        progres = ttk.LabelFrame(app, text="Status", height='85', width='290')
        progres.grid(row=1, column=1, padx=5, pady=0, ipadx=50, ipady=0, sticky='W')
        procesing_line = ttk.Label(progres, text=f"Zakończono")
        procesing_line.place(x=35, y=20)
        procesing_line.update()
        print(final_df)
        final_df.to_excel(f"LOG_FOLDER{todaydate}.xlsx", index=False)


app = Tk()
app.geometry('620x290')
style = ThemedStyle(app)
style.set_theme("arc")
# sv_ttk.set_theme("dark")
app.configure(background='#f5f6f7')
app.resizable(True, True)
app.title("Dodawanie Dokumentow & Zdjec do Step")

# STEP Logowanie Ramka
border = ttk.LabelFrame(app, text="Logowanie do STEP:  ", width='610', height='80')
border.grid(row=0, column=0, columnspan=2, padx=5, pady=0, ipadx=0, ipady=0, sticky='EW')

login = StringVar()
label_step_login = ttk.Label(border, text="Login", width=5, font=("bold", 10))
label_step_login.place(x=10, y=5)
loginEntry = ttk.Entry(border, textvariable=login, width=90)
loginEntry.place(x=50, y=5)

password = StringVar()
label_step_password = ttk.Label(border, text="Hasło", width=5, font=("bold", 10))
label_step_password.place(x=10, y=35)
passwordEntry = ttk.Entry(border, textvariable=password, show='*', width=90)
passwordEntry.place(x=50, y=35)
# Step Logowanie koniec

reference_dictionary = {
    "Zdjęcie Główne": 'Product%20Image',
    "Zdjęcie Dodatkowe": 'Product%20Image%20further',
    "Zdjęcie Wymiarowania": 'MeasurementDrawing',
    "Zdjęcie Techniczne": 'Technical%20Drawing',
    "Zdjęcie Główne Polska": 'LocalPrimaryImage',
    "Dokumenty": 'PL_Documents',
    "Certyfikaty": 'PL_Certificates_Reference'
}

# Ramka wgrywanie select
borderSelect = ttk.LabelFrame(app, text="Wybór danych:  ", width='200', height='200')
borderSelect.grid(row=1, column=0, padx=5, pady=0, ipadx=0, ipady=0, sticky='W')

# Przycisk Wyboru
add_btn = ttk.Button(borderSelect, text="Plik z danymi", command=get_file)
add_btn.place(x=10, y=140)

ratio = IntVar()
Radiobutton = ttk.Radiobutton(borderSelect, text='Wgrywanie z linku', variable=ratio, value=1)
Radiobutton.place(x=10, y=5)
Radiobutton2 = ttk.Radiobutton(borderSelect, text='Wgrywanie z dysku', variable=ratio, value=2)
Radiobutton2.place(x=10, y=50)

add_btn2 = ttk.Button(borderSelect, text="Folder z zdjeciami", width=15, command=get_file_local)
add_btn2.place(x=10, y=95)

# Zatwierdzenie

border4 = ttk.LabelFrame(app, border=0, width='400', height='200')
border4.grid(row=1, column=1, padx=0, pady=0, ipadx=0, ipady=0, sticky='W')
ttk.Button(border4, text='Wgraj do STEP', command=start).place(x=10, y=10)
ttk.Label(border4, text="Created by Adrian Minecki", font=("Arial", 7)).place(x=10, y=168)
ttk.Label(border4, text="ver. " + wersja_prog, font=("Arial", 6)).place(x=130, y=169)

app.mainloop()
