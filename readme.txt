Support Updater Tool v1.4 --8/9/23
Written in Python 3.9.7

SupportUpdater.exe
Program that will detect and update out-of-date subdirectories by comparing a selected source dir to the selected destination dir.
User can select which subdirectories to update or update all at once.
Detects "out-of-date" files by comparing modify times and data sizes.
Uses zip version of support subdirs when possible. Copies zip to temporary local folder and copies files from there. Improves performance. All temporary files are deleted when done.
User can preview to find out which folders in their local support folder are out-of-date before syncing.
User can empty the selected destination directory.
*** Tool requires Administrative Privileges to avoid permission errors when accessing files.
DOES NOT DELETE UNIQUE FILES ON YOUR LOCAL SUPPORT. ANY FILES/FOLDERS THAT ARE NEW OR NEWLY UPDATED ON THE LOCAL FOLDER ARE NOT CHANGED

SupportUpdater.py
Sourced code for the exe. 

nxfemap_app.ico
Femap logo icon. Must be stored in same location as .exe file to be found and used. 


TO EDIT/UPDATE:
Make coding changes to the .py file. To compile into an exe, download pyinstaller (pip install pyinstaller). Open cmd, cd into folder containing .py. 
Run command "pyinstaller --onefile SupportUpdater.py". This will create an .exe in a folder "dist" within original folder. This is your newly compiled .exe.



***DISCALIMER***
Due to the nature of reading mass amounts of file data from a network file, exacerbated by use over a VPN, this tool may take excessive amounts of time.
Non-responsiveness is expected, the script will still be running, just at a rate slower than the native Windows "responsiveness" detection threshold.
To enable debug console, comment out:
	hide = win32gui.GetForegroundWindow()
	win32gui.ShowWindow(hide, win32con.SW_HIDE)
from the main method.

Tyler Grover (tyler.grover@siemens.com / tkg36@drexel.edu)
Siemens AG

