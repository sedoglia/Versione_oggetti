#!/usr/bin/env python3
"""
Versione Python FINALE con messaggio di completamento in modalità interattiva
"""

import argparse
import os
import sys
import csv
import datetime
import multiprocessing
import concurrent.futures
from pathlib import Path
from typing import List, Dict, Optional, Set, Tuple
import ctypes
from ctypes import wintypes, byref, create_string_buffer, c_void_p, c_uint, POINTER, Structure
import struct
import tkinter as tk
from tkinter import filedialog, messagebox

# Definizione delle strutture Windows per accedere alle informazioni di versione
class VS_FIXEDFILEINFO(Structure):
    _fields_ = [
        ('dwSignature', wintypes.DWORD),
        ('dwStrucVersion', wintypes.DWORD),
        ('dwFileVersionMS', wintypes.DWORD),  
        ('dwFileVersionLS', wintypes.DWORD),
        ('dwProductVersionMS', wintypes.DWORD),
        ('dwProductVersionLS', wintypes.DWORD),
        ('dwFileFlagsMask', wintypes.DWORD),
        ('dwFileFlags', wintypes.DWORD),
        ('dwFileOS', wintypes.DWORD),
        ('dwFileType', wintypes.DWORD),
        ('dwFileSubtype', wintypes.DWORD),
        ('dwFileDateMS', wintypes.DWORD),
        ('dwFileDateLS', wintypes.DWORD),
    ]

# Caricamento delle DLL di Windows
try:
    version_dll = ctypes.windll.version

    # Definizione delle funzioni API
    GetFileVersionInfoSizeW = version_dll.GetFileVersionInfoSizeW
    GetFileVersionInfoSizeW.argtypes = [wintypes.LPCWSTR, POINTER(wintypes.DWORD)]
    GetFileVersionInfoSizeW.restype = wintypes.DWORD

    GetFileVersionInfoW = version_dll.GetFileVersionInfoW
    GetFileVersionInfoW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, c_void_p]
    GetFileVersionInfoW.restype = wintypes.BOOL

    VerQueryValueW = version_dll.VerQueryValueW
    VerQueryValueW.argtypes = [c_void_p, wintypes.LPCWSTR, POINTER(c_void_p), POINTER(c_uint)]
    VerQueryValueW.restype = wintypes.BOOL

    WINDOWS_API_AVAILABLE = True
except:
    WINDOWS_API_AVAILABLE = False

# Variabile globale per tracciare la modalità interattiva
INTERACTIVE_MODE = False

def get_folder_path(description="Seleziona la cartella da analizzare", initial_directory=""):
    """
    Mostra un dialog per la selezione della cartella
    Equivalente alla funzione Get-FolderPath di PowerShell
    """
    global INTERACTIVE_MODE
    INTERACTIVE_MODE = True  # Se viene chiamata questa funzione, siamo in modalità interattiva

    try:
        root = tk.Tk()
        root.withdraw()  # Nasconde la finestra principale

        folder_path = filedialog.askdirectory(
            title=description,
            initialdir=initial_directory if initial_directory and os.path.exists(initial_directory) else None
        )

        root.destroy()
        return folder_path if folder_path else None
    except Exception as e:
        print(f"Errore nella selezione della cartella: {e}")
        return None

def show_completion_message():
    """
    Mostra messaggio di completamento in modalità interattiva
    Equivalente al PowerShell:
    [System.Windows.Forms.MessageBox]::Show("Esecuzione terminata.", "VERSIONE_OGGETTI", ...)
    """
    try:
        # Crea una finestra Tkinter temporanea per il messaggio
        root = tk.Tk()
        root.withdraw()  # Nasconde la finestra principale

        # Mostra il messaggio
        messagebox.showinfo(
            title="VERSIONE_OGGETTI",
            message="Esecuzione terminata."
        )

        root.destroy()
    except Exception:
        # Se c'è un errore, ignora silenziosamente (come nel PowerShell)
        pass

def get_script_path():
    """
    Ottiene il percorso dello script corrente
    Equivalente alla logica PowerShell per ottenere $strScriptPath
    """
    if getattr(sys, 'frozen', False):
        # Script compilato con PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Script Python normale
        return os.path.dirname(os.path.abspath(__file__))

def create_semaphore_file(script_path):
    """
    Crea il file di semaforo equivalente al PowerShell
    """
    semaphore_path = os.path.join(script_path, "999_VOGG.TXT")
    semaphore_info = [
        str(os.getpid()),
        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        os.environ.get('COMPUTERNAME', os.environ.get('HOSTNAME', 'Unknown')),
        os.environ.get('USERNAME', os.environ.get('USER', 'Unknown'))
    ]

    try:
        with open(semaphore_path, 'w', encoding='utf-8') as f:
            for info in semaphore_info:
                f.write(info + '\n')
        return semaphore_path
    except Exception as e:
        print(f"Errore nella creazione del file di semaforo: {e}")
        return None

def remove_semaphore_file(semaphore_path):
    """
    Rimuove il file di semaforo
    """
    try:
        if semaphore_path and os.path.exists(semaphore_path):
            os.remove(semaphore_path)
    except Exception:
        pass  # Ignora errori nella rimozione

def get_all_available_translations(buffer):
    """
    Ottiene tutte le traduzioni disponibili nel file
    """
    translations = []

    try:
        ptr = c_void_p()
        length = c_uint()

        if VerQueryValueW(buffer, "\\VarFileInfo\\Translation", byref(ptr), byref(length)):
            if length.value >= 4:
                for i in range(0, length.value, 4):
                    try:
                        lang = struct.unpack('<H', ctypes.string_at(ptr.value + i, 2))[0]
                        cp = struct.unpack('<H', ctypes.string_at(ptr.value + i + 2, 2))[0]
                        translations.append(f"{lang:04X}{cp:04X}")
                    except:
                        continue
    except:
        pass

    # Se non trovate traduzioni, usa quelle comuni + altre varianti
    if not translations:
        translations = [
            '040904B0',  # English (US) - Unicode
            '040904E4',  # English (US) - Windows-1252
            '041004B0',  # Italian - Unicode  
            '041004E4',  # Italian - Windows-1252
            '040704B0',  # German - Unicode
            '040C04B0',  # French - Unicode
            '080904B0',  # English (UK) - Unicode
            '000004B0',  # Neutral - Unicode
            '000004E4'   # Neutral - Windows-1252
        ]

    return translations

def get_string_from_version_info(buffer, key, translations):
    """
    Cerca una stringa nelle informazioni di versione usando tutte le traduzioni disponibili
    """
    for translation in translations:
        try:
            query_str = f"\\StringFileInfo\\{translation}\\{key}"
            s_ptr = c_void_p()
            s_len = c_uint()

            if VerQueryValueW(buffer, query_str, byref(s_ptr), byref(s_len)):
                if s_len.value > 0:
                    try:
                        value = ctypes.wstring_at(s_ptr.value)
                        if value and value.strip():
                            return value.strip()
                    except:
                        continue
        except:
            continue

    return ""

def get_version_info_local(file_path: str, key: str) -> Optional[Dict[str, str]]:
    """
    REPLICA ESATTA della funzione Get-VersionInfoLocal del PowerShell
    Cerca specificatamente la chiave richiesta nelle informazioni di versione
    """
    if not WINDOWS_API_AVAILABLE:
        return None

    try:
        handle = wintypes.DWORD(0)
        size = GetFileVersionInfoSizeW(file_path, byref(handle))

        if size <= 0:
            return None

        buffer = create_string_buffer(size)

        if not GetFileVersionInfoW(file_path, 0, size, buffer):
            return None

        # Ottieni tutte le traduzioni disponibili
        translations = get_all_available_translations(buffer)

        # Cerca il valore per la chiave specifica
        value = get_string_from_version_info(buffer, key, translations)

        if value:
            return {
                'Value': value,
                'Translation': 'Found'
            }

        return None

    except Exception:
        return None

def get_file_version_info(file_path: str) -> Tuple[str, str]:
    """
    Ottiene ProductName e FileVersion usando VersionInfo
    REPLICA ESATTA di: (Get-Item $FullName).VersionInfo.ProductName, (Get-Item $FullName).VersionInfo.FileVersion
    CON FOCUS SPECIALE SU ProductName che non veniva estratto correttamente
    """
    product_name = ""
    file_version = ""

    try:
        # Prova prima con le API native (più controllo)
        if WINDOWS_API_AVAILABLE:
            handle = wintypes.DWORD(0)
            size = GetFileVersionInfoSizeW(file_path, byref(handle))

            if size > 0:
                buffer = create_string_buffer(size)

                if GetFileVersionInfoW(file_path, 0, size, buffer):
                    # Ottieni tutte le traduzioni disponibili
                    translations = get_all_available_translations(buffer)

                    # FileVersion dalla struttura VS_FIXEDFILEINFO
                    try:
                        ptr = c_void_p()
                        length = c_uint()

                        if VerQueryValueW(buffer, "\\", byref(ptr), byref(length)):
                            if length.value >= ctypes.sizeof(VS_FIXEDFILEINFO):
                                fixed_info = ctypes.cast(ptr, POINTER(VS_FIXEDFILEINFO)).contents

                                major = (fixed_info.dwFileVersionMS >> 16) & 0xFFFF
                                minor = fixed_info.dwFileVersionMS & 0xFFFF
                                build = (fixed_info.dwFileVersionLS >> 16) & 0xFFFF
                                revision = fixed_info.dwFileVersionLS & 0xFFFF

                                file_version = f"{major}.{minor}.{build}.{revision}"
                    except:
                        pass

                    # ProductName con ricerca estensiva in tutte le traduzioni
                    product_name = get_string_from_version_info(buffer, "ProductName", translations)

                    # Se ProductName è vuoto, prova altre chiavi correlate
                    if not product_name:
                        # Prova FileDescription come alternativa
                        product_name = get_string_from_version_info(buffer, "FileDescription", translations)

                    # Se ancora vuoto, prova InternalName
                    if not product_name:
                        product_name = get_string_from_version_info(buffer, "InternalName", translations)

                    # Se ancora vuoto, prova OriginalFilename senza estensione
                    if not product_name:
                        original_name = get_string_from_version_info(buffer, "OriginalFilename", translations)
                        if original_name:
                            # Rimuovi estensione
                            product_name = os.path.splitext(original_name)[0]

        # Fallback con pywin32 se le API native non hanno funzionato
        if (not product_name or not file_version):
            try:
                import win32api

                # FileVersion se non ancora ottenuto
                if not file_version:
                    try:
                        version_info = win32api.GetFileVersionInfo(file_path, "\\")
                        ms = version_info['FileVersionMS']
                        ls = version_info['FileVersionLS']
                        file_version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
                    except:
                        pass

                # ProductName se non ancora ottenuto - prova varie codepage
                if not product_name:
                    codepages = ['040904B0', '040904E4', '041004B0', '041004E4', '000004B0', '040704B0', '040C04B0']

                    for cp in codepages:
                        try:
                            product_name = win32api.GetFileVersionInfo(file_path, f"\\StringFileInfo\\{cp}\\ProductName")
                            if product_name and product_name.strip():
                                product_name = product_name.strip()
                                break
                        except:
                            continue

                    # Se ProductName ancora vuoto, prova FileDescription
                    if not product_name:
                        for cp in codepages:
                            try:
                                product_name = win32api.GetFileVersionInfo(file_path, f"\\StringFileInfo\\{cp}\\FileDescription")
                                if product_name and product_name.strip():
                                    product_name = product_name.strip()
                                    break
                            except:
                                continue

                    # Se ancora vuoto, prova InternalName
                    if not product_name:
                        for cp in codepages:
                            try:
                                product_name = win32api.GetFileVersionInfo(file_path, f"\\StringFileInfo\\{cp}\\InternalName")
                                if product_name and product_name.strip():
                                    product_name = product_name.strip()
                                    break
                            except:
                                continue

            except ImportError:
                pass

    except Exception:
        pass

    # Se ProductName è ancora vuoto, usa il nome del file senza estensione
    if not product_name:
        product_name = os.path.splitext(os.path.basename(file_path))[0]

    # Se FileVersion è vuoto, usa "0.0.0.0"
    if not file_version:
        file_version = "0.0.0.0"

    return product_name, file_version

def process_file(file_info: Tuple[str, str]) -> Optional[Dict[str, any]]:
    """
    Processa un singolo file per estrarre le informazioni di versione
    REPLICA ESATTA della logica PowerShell con ProductName corretto
    """
    full_path, key = file_info

    try:
        # Ottieni informazioni del file
        file_stat = os.stat(full_path)
        last_modified = datetime.datetime.fromtimestamp(file_stat.st_mtime)

        # STEP 1: Prova a ottenere il valore per la chiave specifica
        # Equivalente a: $valueInfo = Try-GetVersionObject $FullName $Key
        value_info = get_version_info_local(full_path, key)

        # STEP 2: Applica la logica condizionale esatta del PowerShell
        if value_info and value_info.get('Value'):
            # Se trovato valore per la chiave, usalo
            package_value = value_info['Value']
        else:
            # FALLBACK: Usa ProductName - FileVersion (ORA CORRETTO!)
            # Equivalente a: "{0} - {1}" -f (Get-Item $FullName).VersionInfo.ProductName, (Get-Item $FullName).VersionInfo.FileVersion
            product_name, file_version = get_file_version_info(full_path)
            package_value = f"{product_name} - {file_version}"

        return {
            'PercorsoCompleto': full_path,
            'Package': package_value,
            'DataOra': last_modified.strftime('%Y-%m-%d %H:%M:%S'),
            'Dimensione': file_stat.st_size
        }

    except Exception as e:
        return None

def scan_files(root_paths: List[str]) -> List[str]:
    """
    Scansiona le cartelle per trovare file .exe e .dll
    Equivalente al blocco "Scansione file" di PowerShell
    """
    all_files = []
    seen_paths = set()

    for root_path in root_paths:
        try:
            # Risolvi il percorso
            resolved_path = os.path.abspath(root_path)

            if os.path.exists(resolved_path):
                for root, dirs, files in os.walk(resolved_path):
                    for file in files:
                        if file.lower().endswith(('.exe', '.dll')):
                            full_path = os.path.join(root, file)
                            if full_path not in seen_paths:
                                seen_paths.add(full_path)
                                all_files.append(full_path)
        except Exception as e:
            continue  # Ignora errori e continua con il prossimo percorso

    return all_files

def main():
    """
    Funzione principale - equivalente al main script PowerShell
    """
    global INTERACTIVE_MODE

    parser = argparse.ArgumentParser(description='Analizza file .exe e .dll per estrarre informazioni di versione')
    parser.add_argument('--root-path', nargs='*', help='Percorsi delle cartelle da analizzare')
    parser.add_argument('--output-csv', help='Percorso del file CSV di output')
    parser.add_argument('--delimiter', default=';', help='Delimitatore CSV (default: ;)')
    parser.add_argument('--key', default='Package', help='Chiave per estrazione versione (default: Package)')
    parser.add_argument('--max-threads', type=int, help='Numero massimo di thread')

    args = parser.parse_args()

    # Gestione percorsi root
    root_paths = args.root_path if args.root_path else []

    if not root_paths:
        selected_path = get_folder_path("Seleziona la cartella da analizzare")
        if selected_path:
            root_paths = [selected_path]
            # Verifica percorsi UNC (Windows)
            if selected_path.startswith('\\\\'):
                if not os.path.exists(selected_path):
                    print("Avviso: Il percorso UNC potrebbe non essere accessibile. Continuando comunque...")
        else:
            print("Nessuna cartella selezionata. Uscita.")
            return

    # Ottieni percorso script
    script_path = get_script_path()

    # Imposta percorso output CSV
    output_csv_path = args.output_csv
    if not output_csv_path:
        output_csv_path = os.path.join(script_path, "VERSIONE_OGGETTI.CSV")

    # Crea file di semaforo
    semaphore_path = create_semaphore_file(script_path)

    try:
        # Scansiona i file
        all_files = scan_files(root_paths)

        if not all_files:
            if INTERACTIVE_MODE:
                print("Nessun file .exe o .dll trovato.")
            return

        # Calcola numero thread
        THREADS_MIN = 1
        THREADS_CAP = 32
        logical_processors = multiprocessing.cpu_count()
        computed_threads = min(THREADS_CAP, max(THREADS_MIN, 2 * logical_processors))

        max_threads = args.max_threads if args.max_threads else computed_threads
        max_threads = min(THREADS_CAP, max(THREADS_MIN, max_threads))

        if INTERACTIVE_MODE:
            print(f"Elaborazione di {len(all_files)} file con {max_threads} thread...")

        # Processa i file in parallelo
        file_infos = [(file_path, args.key) for file_path in all_files]
        results = []

        with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
            future_to_file = {executor.submit(process_file, file_info): file_info for file_info in file_infos}

            for future in concurrent.futures.as_completed(future_to_file):
                result = future.result()
                if result:
                    results.append(result)

        # Scrivi il CSV
        if results:
            with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['PercorsoCompleto', 'Package', 'DataOra', 'Dimensione']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=args.delimiter)

                writer.writeheader()
                for row in results:
                    writer.writerow(row)

            if INTERACTIVE_MODE:
                print(f"Creato: {output_csv_path} (righe: {len(results)})")
        else:
            if INTERACTIVE_MODE:
                print("Nessun dato esportato.")

    finally:
        # Rimuovi il file di semaforo
        remove_semaphore_file(semaphore_path)

        # >>> Messaggio finale SOLO in modalità interattiva (nessun output in modalità silente/CLI)
        if INTERACTIVE_MODE:
            show_completion_message()

if __name__ == "__main__":
    main()
