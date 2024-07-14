# ver 1.1.1
# internal libraries
from os import rename
import sys
import subprocess
import pkg_resources
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)

# need pip install setuptools for pkg_resources
# check for required packages - install if missing
required = {'snowflake-connector-python', 'pandas', 'numpy', 'openpyxl', 'tk', 'pywin32', 'requests', 'python-certifi-win32', 'pyodbc'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed

if missing:
	python = sys.executable
	subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)
	ctypes.windll.user32.MessageBoxW(0, "SETUP COMPLETE: RESTART THE PROGRAM", "LIBRARY INSTALL", 0x0 | 0x40)
	sys.exit()

		
# external libraries
from TP2STK_backend import *
import tkinter as tk
from tkinter import messagebox


def userInterface():

	root = tk.Tk()
	root.title('TP2STK')
	root.iconbitmap('DEVELOPER_FILES/icon_box.ico')

	bttn_split = tk.Button(
		root,
		text='CREATE TEMPLATE FILES',
		font=('Arial bold', 11), 
		width=30, 
		height=2,
		borderwidth=4,
		bg='#3E4149',
		fg='white',
		command=confirmSplit
		)
	bttn_split.pack(ipadx=5, ipady=5, expand=True)

	bttn_SMTP = tk.Button(
		root, 
		text='SEND SMTP EMAIL', 
		font=('Arial bold', 11), 
		width=30, 
		height=2,
		borderwidth=4,
		bg='#3E4149',
		fg='white', 
		command=confirmSMTP
		)

	bttn_SMTP.pack(ipadx=5, ipady=5, expand=True)

	bttn_outlook = tk.Button(
		root, 
		text='SEND OUTLOOK EMAIL', 
		font=('Arial bold', 11), 
		width=30, 
		height=2,
		borderwidth=4,
		bg='#3E4149',
		fg='white', 
		command=confirmOutlook
		)
	bttn_outlook.pack(ipadx=5, ipady=5, expand=True)

	bttn_stitch = tk.Button(
		root, 
		text='IMPORT SUPPLIER FEEDBACK', 
		font=('Arial bold', 11), 
		width=30, 
		height=2,
		borderwidth=4,
		bg='#3E4149',
		fg='white', 
		command=confirmStitch
		)
	bttn_stitch.pack(ipadx=5, ipady=5, expand=True)

	root.mainloop()


def confirmSMTP():

	answer = messagebox.askyesno(title='SEND SUPPLIER EMAIL', message='ARE YOU SURE YOU WANT TO SEND SMTP EMAIL?')
	if answer:
		sendSMTP()


def confirmOutlook():

	answer = messagebox.askyesno(title='SEND SUPPLIER EMAIL', message='ARE YOU SURE YOU WANT TO SEND OUTLOOK EMAIL?')
	if answer:
		sendOutlook()


def confirmSplit():

	# check if files already exist
	files = Path('SPLIT_FILES').glob('*.xlsx')
	file_list = [path.abspath(filepath) for filepath in files]
	
	if file_list:
		answer = messagebox.askyesno(title='SPLIT FILES EXIST', message='ARE YOU SURE YOU WANT TO REPLACE CURRENT FILES?')
		if answer:
			splitFiles()
	else:
		splitFiles()


def confirmStitch():

	# check if files exist
	files = Path('IMPORT_FILES').glob('*.xlsx')
	file_list = [path.abspath(filepath) for filepath in files]
	
	if not file_list:
		messagebox.showwarning(title='IMPORT FILES ERROR', message='NO FILES TO IMPORT - MAKE SURE THERE ARE FILES IN IMPORT_FILES FOLDER')
		return
	else:
		stitchFiles()

if __name__ == "__main__":
	userInterface()