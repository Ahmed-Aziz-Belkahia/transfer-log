#you can't run this program from "main.py" you must download it from here               |
#"https://drive.google.com/file/d/1QJCXjwixox990A1eff92958lxQuLLv9i/view?usp=sharing"  _|

import psutil
import os, sys

processes = 0

for i in psutil.process_iter():
    if "Transfer log.exe" == i.name():
        processes += 1
if processes > 1:
    sys.exit(-1)

from functools import partial
from posixpath import splitext
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import asksaveasfile
from PIL import Image
from win32com.shell import shell, shellcon

import sqlite3 as sql
import time
import customtkinter
import csv
import pystray
import win32api
import json
import os.path
import shutil

SPATH = ""

with open(os.getenv('APPDATA')+"\Darkonex\APPDIR.json") as file:
    SPATH = json.load(file)    

def make_shortcut():
    src = f"{SPATH}\Transfer log startup.lnk"
    dist_path = shell.SHGetFolderPath (0, shellcon.CSIDL_STARTUP, 0, 0)+"\Transfer log startup.lnk"
    shutil.copyfile(src, dist_path)

#creating window
main = customtkinter.CTk()
main.title("transfer log")
main.iconbitmap(f"{SPATH}/icon white.ico")
main.geometry(f"{780}x{655}")
main.minsize(780, 300)

#initialize variables
hide_TMP = True
close = False
Tmp_Exts = ""
From_Dir = None

#getting tmp extensions
with open(f"{SPATH}/tmp extensions.json", "r") as file:
    Tmp_Exts = json.load(file)

#create a database and a cursor
conn = sql.connect(f"{SPATH}/Log.db", check_same_thread=False)
cursor = conn.cursor()
cursor.execute("""CREATE TABLE IF NOT EXISTS transfer_log(
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    File_name VARCHAR(255),
    From_Dir VARCHAR(255),
    To_Dir VARCHAR(255))
    """)

#indexing File_name and To_Dir and From_Dir
cursor.execute("CREATE INDEX IF NOT EXISTS File_name_index ON transfer_log (File_name)")
cursor.execute("CREATE INDEX IF NOT EXISTS To_Dir_index ON transfer_log (To_Dir)")
cursor.execute("CREATE INDEX IF NOT EXISTS From_Dir ON transfer_log (From_Dir)")

#creating event handler
my_event_handler = PatternMatchingEventHandler(["*"], None, False, False)

#settings tkinter variables
Run_In_Background = IntVar()
Run_On_Startup = IntVar()
Tray_icon_status = IntVar()
log_dup = IntVar()
log_tmp = IntVar()
search_keyword = StringVar()

#open or create settings file
if os.path.exists(f"{SPATH}/Settings.json"):
    with open(f"{SPATH}/Settings.json") as file:
        settings = json.load(file)
        Run_In_Background.set(settings["Run_In_Background"])
        Run_On_Startup.set(settings["Run_On_Startup"])
        log_dup.set(settings["log_dup"])
        log_tmp.set(settings["log_tmp"])
else:
    Settings = {"Run_In_Background" : 0, "log_dup" : 1, "log_tmp" : 0, "Run_On_Startup" : 0}
    with open(f"{SPATH}/Settings.json", "w") as file:
        json.dump(Settings, file, indent=4)

def Run_on_startup_check():
    Update_Settings()
    if Run_On_Startup.get() == 1:
        while True:
            try:
                make_shortcut()
                break
            except:
                os.remove(f"{shell.SHGetFolderPath (0, shellcon.CSIDL_STARTUP, 0, 0)}\Transfer log startup.lnk")
    else:
        try:
            os.remove(f"{shell.SHGetFolderPath (0, shellcon.CSIDL_STARTUP, 0, 0)}\Transfer log startup.lnk")
        except:pass

#export log to csv file
def Export():
    file = asksaveasfile(filetypes=(("csv file", "*.csv"),), title="Export file", initialfile="log",  defaultextension = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
    writer = csv.writer(file)
    for row in cursor.execute("SELECT * FROM transfer_log").fetchall():
        writer.writerow(row)

#when i close the window
def raiserr():
    if Run_In_Background.get() == 1:
        main.withdraw()
        CreateTray()
        Tray.run()
    else:
        main.destroy()
        globals()["close"] = True

#clear log and treeview table
def Clear():
    cursor.execute("DELETE FROM transfer_log")
    conn.commit()
    table.delete(*table.get_children())

#when i press open from tray icon
def show():
    Tray.stop()
    main.deiconify()

#when i exit from manu bar or tray icon
def Exit():
    try:Tray.stop()
    except:pass
    globals()["close"] = True
    try:main.destroy()
    except:sys.exit(-1)

#create a tray icon
def CreateTray():
        icon = Image.open(f"{SPATH}/icon white.ico")
        global Tray
        Tray = pystray.Icon("transfer log", icon, title="transfer log", menu=pystray.Menu(
            pystray.MenuItem("Open", show),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Export", Export),
            pystray.MenuItem("Clear log", Clear),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", Exit),))

#function to call everytime i change something from settings to submit it to settings.json
def Update_Settings():
    settings = {"Run_In_Background" : Run_In_Background.get(), "log_dup" : log_dup.get(), "log_tmp" : log_tmp.get(), "Run_On_Startup" : Run_On_Startup.get()}
    Run_In_Background.set(settings["Run_In_Background"])
    Run_On_Startup.set(settings["Run_On_Startup"])
    log_dup.set(settings["log_dup"])
    log_tmp.set(settings["log_tmp"])
    with open(f"{SPATH}/Settings.json", "w") as file:
        json.dump(settings, file, indent=4)

#creating menu bar
menubar = Menu(main)
main.config(menu= menubar)

menu_file = Menu(menubar, tearoff=0)
menu_edit = Menu(menubar, tearoff=0)
menubar.add_cascade(menu=menu_file, label='File')
menubar.add_cascade(menu=menu_edit, label='Edit')

menu_file.add_command(label='Export', command=Export)
menu_file.add_command(label="Clear log", command=Clear)
menu_file.add_separator()
menu_file.add_command(label='Exit', command=Exit)

if Run_In_Background.get() == 1 and '-startup' in sys.argv:
    raiserr()

try:
    menu_edit.add_checkbutton(label="Run in background", onvalue=1, offvalue=0, variable=Run_In_Background, command=Update_Settings)
    menu_edit.add_checkbutton(label="Run on startup", onvalue=1, offvalue=0, variable=Run_On_Startup, command=Run_on_startup_check)
    menu_edit.add_checkbutton(label="log duplicate files", onvalue=1, offvalue=0, variable=log_dup, command=Update_Settings)
    menu_edit.add_checkbutton(label="log .tmp files", onvalue=1, offvalue=0, variable=log_tmp, command=Update_Settings)
except: sys.exit()

style = ttk.Style()
style.theme_use("clam")

#styling our treeview
style.configure("Treeview",
    background="#212325",
    foreground="white",
    rowheight=25,
    fieldbackground="#212325",
    borderwidth=0,
)
style.configure('Treeview.Heading', relief="flat", background='#212325', foreground='white', bordercolor="#212325", borderwidth=0, )
style.map( 'Treeview' ,background = [( ' selected ' , ' gray' )])
columns = ['File', 'From', 'To']
table = ttk.Treeview(main, show='headings', selectmode='none', height=30)
#creating columns
table['columns'] = columns
table.column('File', anchor="w", width=50)
table.column("From", anchor="w", width=280)
table.column('To', anchor="w", width=280)

file_mode = 0
from_mode = 0
to_mode = 0

#on heading click
def sort(column):
    if column == 'File':
        globals()["file_mode"] +=1
        if globals()["file_mode"] == 2:
            globals()["file_mode"] = 0
            rows = cursor.execute("SELECT * FROM transfer_log ORDER BY File_name ASC").fetchall()
        else:
            rows = cursor.execute("SELECT * FROM transfer_log ORDER BY File_name DESC").fetchall()
    elif column == 'From':
        globals()["from_mode"] +=1
        if globals()["from_mode"] == 2:
            globals()["from_mode"] = 0
            rows = cursor.execute("SELECT * FROM transfer_log ORDER BY From_Dir ASC").fetchall()
        else:rows = cursor.execute("SELECT * FROM transfer_log ORDER BY From_Dir DESC").fetchall()
    elif column == 'To':
        globals()["to_mode"] +=1
        if globals()["to_mode"] == 2:
            globals()["to_mode"] = 0
            rows = cursor.execute("SELECT * FROM transfer_log ORDER BY To_Dir ASC").fetchall()
        else:
            rows = cursor.execute("SELECT * FROM transfer_log ORDER BY To_Dir DESC").fetchall()

    table.delete(*table.get_children())
    for i in range(cursor.execute("SELECT COUNT(*) FROM transfer_log").fetchone()[0]):
        row = rows[i]
        table.insert(parent="", index="end", iid=row[0], text="parent", values=(row[1], row[2], row[3]))        

for column in columns:
    table.heading(column, text=column, command=partial(sort, column))
rows = cursor.execute("SELECT * FROM transfer_log").fetchall()
def inser_db_to_table():
    for i in range(cursor.execute("SELECT COUNT(*) FROM transfer_log").fetchone()[0]):
        row = rows[i]
        try:
            table.insert(parent="", index="end", iid=row[0], text="parent", values=(row[1], row[2], row[3]))
        except:continue
inser_db_to_table()

def search(event):
    keyword = f"{search_keyword.get()}%"
    if len(keyword)-1 == 0:
        inser_db_to_table()
        return
    rows = cursor.execute("SELECT * FROM transfer_log WHERE File_name LIKE ?", [keyword]).fetchall()
    for row in table.get_children():
        table.delete(row)
    for i in range(len(rows)):
        row = rows[i]
        table.insert(parent="", index="end", iid=row[0], text="parent", values=(row[1], row[2], row[3]))

#create search bar
search_bar = customtkinter.CTkEntry(main, placeholder_text="search", border_width=2, textvariable=search_keyword)
search_bar.bind("<KeyRelease>", search)
search_bar.pack(fill="x")

table.pack(expand=True, fill="both")

#event functions
def on_deleted(event):
    globals()["From_Dir"] = event.src_path

def on_created(event):
    CurrentDir = event.src_path
    if From_Dir == None:
        return

    old_file_name = os.path.basename(From_Dir)
    old_file_name_ext = splitext(old_file_name)[1].lower()

    Curent_file_name = os.path.basename(CurrentDir)

    if log_tmp.get() == 0 and old_file_name_ext in Tmp_Exts or From_Dir == CurrentDir:
        return
    if old_file_name == Curent_file_name:
        if log_dup.get() == 0:
            file_name = os.path.basename(From_Dir)
            cursor.execute("DELETE FROM transfer_log WHERE ID = (SELECT ID FROM transfer_log WHERE File_name = ? AND To_Dir = ? ORDER BY ID DESC LIMIT 1)", (file_name, From_Dir))
            for row in table.get_children():
                table.delete(row)
            for i in range(cursor.execute("SELECT COUNT(*) FROM transfer_log").fetchone()[0]):
                row = cursor.execute("SELECT * FROM transfer_log").fetchall()[i]
                table.insert(parent="", index="end", iid=row[0], text="parent", values=(row[1], row[2], row[3]))
        cursor.execute("INSERT INTO transfer_log(File_name, From_Dir, To_Dir) VALUES(?, ?, ?)", (old_file_name, From_Dir, CurrentDir))
        conn.commit()
        table.insert(parent="", index="end", text="parent", values=(old_file_name, From_Dir, CurrentDir))

my_event_handler.on_created = on_created
my_event_handler.on_deleted = on_deleted


drives = win32api.GetLogicalDriveStrings()
paths = drives.split('\000')[:-1]

observer = Observer()

for path in paths:
    targetPath = path
    observer.schedule(my_event_handler, targetPath, recursive=True)
observer.start()
Run_on_startup_check()

main.protocol("WM_DELETE_WINDOW", raiserr)
main.mainloop()

while close == False:
    try:
        if close == True:
            raise KeyboardInterrupt
        time.sleep(1)
    except KeyboardInterrupt:
        conn.close()
        observer.stop()