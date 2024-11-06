#Imports
import wx
import ctypes, os
import filecmp
import shutil
import traceback
import sys, subprocess
from zipfile import ZipFile
import win32com.client as win32
import win32gui, win32con
win32.gencache.is_readonly=False

#Checks for admin privs
def isAdmin():
    try:
        #sys.stdout.write (str(ctypes.windll.shell32.IsUserAnAdmin()==1))
        return ctypes.windll.shell32.IsUserAnAdmin()==1
    except:
        #sys.stdout.write ("FAILURE CHECKING ADMINISTRATOR PRIVILEGES")
        return False



if isAdmin():
    #General inherited class for frames, includes a number of boxsizers and constants
    class MainFrame(wx.Frame):
        def __init__(self, title):
            wx.Frame.__init__(self, None, title=title, size=(700,750))
            self.Bind(wx.EVT_CLOSE, self.OnClose)

            #Constants--------------------------------------------------------------------|
            self.spacer=5
            self.altDirName=''
            self.destDirName=''
            
            #Boxes--------------------------------------------------------------------|
            self.panel = wx.Panel(self)
            self.mainBox = wx.BoxSizer(wx.VERTICAL)
            self.box1 = wx.BoxSizer(wx.HORIZONTAL)
            self.box2 = wx.BoxSizer(wx.HORIZONTAL)
            self.box3 = wx.BoxSizer(wx.HORIZONTAL)
            self.box4 = wx.BoxSizer(wx.HORIZONTAL)
            self.box5 = wx.BoxSizer(wx.HORIZONTAL)
            self.box6 = wx.BoxSizer(wx.HORIZONTAL)
            self.box7 = wx.BoxSizer(wx.HORIZONTAL)

            self.mainBox.Add(self.box1, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=int(.1*self.spacer))
            self.mainBox.Add(self.box2, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=int(.1*self.spacer))
            self.mainBox.Add(self.box3, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=int(.1*self.spacer))
            self.mainBox.Add(self.box4, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=self.spacer*2)
            self.mainBox.Add(self.box5, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=self.spacer)
            self.mainBox.Add(self.box6, proportion=1, flag= wx.ALL | wx.EXPAND, border=int(self.spacer))
            self.mainBox.Add(self.box7, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=self.spacer)

            self.panel.SetSizer(self.mainBox)
            self.panel.Layout()

        #Process for closing window 
        def OnClose(self, event):
            #NEED TO HANDLE HOW THIS ONLY DESTORYS CURRENT WINDOW AND LEAVES THE PROGRAM HANGING
            wx.Exit()
    
    class compareFrame(MainFrame):
        def __init__(self, title):
            super().__init__(title)

            #box1--------------------------------------------------------------------|
            m_text = wx.StaticText(self.panel, -1, "Support Updater Tool", style=wx.ALIGN_RIGHT)
            m_text.SetFont(wx.Font(14, wx.SWISS, wx.NORMAL, wx.BOLD))
            m_text.SetSize(m_text.GetBestSize())
            self.box1.Add(m_text, proportion=1, flag= wx.ALIGN_CENTER, border=self.spacer)

            #box2--------------------------------------------------------------------|
            subbox1=wx.BoxSizer(wx.VERTICAL)
            self.box2.Add(subbox1, flag=wx.ALL | wx.ALIGN_TOP)

            self.box2.AddSpacer(50)

            subbox2=wx.BoxSizer(wx.VERTICAL)
            subbox2.SetMinSize(300,10)
            self.box2.Add(subbox2, flag=wx.ALL | wx.ALIGN_TOP)

            self.dirLabel = wx.StaticText(self.panel, -1, "Directory to pull from")
            subbox1.Add(self.dirLabel, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=int(self.spacer))

            self.directories=self.dDownPopulate()
            self.dirDDown=wx.ComboBox(self.panel, -1, '', choices=self.directories)
            self.dirDDown.SetMinSize((200,10))
            self.dirDDown.Bind(wx.EVT_COMBOBOX, self.updateListBox)
            subbox1.Add(self.dirDDown, 1, wx.ALL | wx.ALIGN_CENTER, self.spacer)

            self.altDirButton=wx.Button(self.panel, id=0, label="Choose Alternate Directory")
            self.altDirButton.Bind(wx.EVT_BUTTON, self.altDir)
            subbox1.Add(self.altDirButton, 1, wx.ALL | wx.ALIGN_CENTER, border=int(self.spacer))

            self.dirLabel = wx.StaticText(self.panel, -1, "Directory to sync to")
            subbox2.Add(self.dirLabel, proportion=1, flag= wx.ALL | wx.ALIGN_CENTER, border=int(self.spacer))

            self.destDirText=wx.TextCtrl(self.panel, name='destDirText')
            self.destDirText.SetMinSize((350,10))
            subbox2.Add(self.destDirText, 1, wx.ALL | wx.ALIGN_CENTER, int(self.spacer))

            self.destDirButton=wx.Button(self.panel, id=0, label="Choose Directory")
            self.destDirButton.Bind(wx.EVT_BUTTON, self.destDir)
            subbox2.Add(self.destDirButton, 1, wx.ALL | wx.ALIGN_CENTER, int(self.spacer))

            '''subbox4=wx.BoxSizer(wx.HORIZONTAL)
            subbox2.Add(subbox4, flag= wx.ALIGN_CENTER)
            
            self.openDestDirButton=wx.BitmapButton(self.panel, id=wx.ID_ADD, bitmap=wx.ID_HELP)
            test=self.openDestDirButton.GetBitmap()
            self.openDestDirButton.SetMinSize((100,100))
            #self.openDestDirButton.SetBitmap(wx.ID_OPEN)
            self.openDestDirButton.Bind(wx.EVT_BUTTON, self.openDestDir)
            subbox4.Add(self.openDestDirButton, 1, wx.ALL , int(self.spacer))'''

            #Box3--------------------------------------------------------------------|
            self.ListBoxText=wx.StaticText(self.panel, label="Choose directories to sync. \nSelect none to sync all directories.", style=wx.ALIGN_CENTER_HORIZONTAL)
            self.ListBoxText.SetSize(self.ListBoxText.GetBestSize())
            self.box3.Add(self.ListBoxText, 1, wx.ALL | wx.ALIGN_CENTER, 0*self.spacer)

            #Box4--------------------------------------------------------------------|
            subbox3=wx.BoxSizer(wx.VERTICAL)
            self.box4.Add(subbox3, wx.ALL | wx.ALIGN_TOP)

            self.subdirChoices=[]
            self.subdirSelect=wx.ListBox(self.panel, -1, choices=self.subdirChoices, style= wx.LB_MULTIPLE | wx.LB_NEEDED_SB, name="CHOOSE")
            self.subdirSelect.SetMinSize((200,300))
            subbox3.Add(self.subdirSelect, 1, wx.ALL | wx.ALIGN_CENTER, 0*self.spacer)

            #Box5--------------------------------------------------------------------|
            self.clearSelectButt=wx.Button(self.panel, -1, "Clear Selection", size=(wx.Size(150,35)))
            self.clearSelectButt.Bind(wx.EVT_BUTTON, self.clearSelect)
            self.box5.Add(self.clearSelectButt, 1, wx.ALL | wx.ALIGN_CENTER, self.spacer)

            self.allSelectButt=wx.Button(self.panel, -1, "Select All", size=(wx.Size(150,35)))
            self.allSelectButt.Bind(wx.EVT_BUTTON, self.allSelect)
            self.box5.Add(self.allSelectButt, 1, wx.ALL | wx.ALIGN_CENTER, self.spacer)

            #Box6--------------------------------------------------------------------|
            self.bigRedButton=wx.Button(self.panel, -1, "Clear Destination Directory")
            self.bigRedButton.SetMinSize(wx.Size(200,50))
            self.bigRedButton.Bind(wx.EVT_BUTTON, self.destruct)
            self.bigRedButton.SetBackgroundColour((250, 0, 0))
            #self.bigRedButton.SetForegroundColour((200,0,0))
            self.bigRedButton.SetSize(self.bigRedButton.GetBestSize())
            self.box7.Add(self.bigRedButton, 1, wx.ALL | wx.ALIGN_BOTTOM, int(self.spacer))

            self.previewButton=wx.Button(self.panel, -1, "Preview Sync")
            self.previewButton.SetMinSize(wx.Size(100,50))
            self.previewButton.Bind(wx.EVT_BUTTON, self.previewSync)
            self.box6.Add(self.previewButton, 1, wx.ALL | wx.ALIGN_BOTTOM, int(self.spacer))

            self.compareButton=wx.Button(self.panel, -1, "Sync directories")
            self.compareButton.SetMinSize(wx.Size(100,50))
            self.compareButton.Bind(wx.EVT_BUTTON, self.syncDirs)
            self.box6.Add(self.compareButton, 1, wx.ALL | wx.ALIGN_BOTTOM, int(self.spacer))

            #Button Default
            ''' 
            self.____=wx.Button(self.panel, -1, "NAME")
            self.____.Bind(wx.EVT_BUTTON, FUNCTION)
            box4.Add(self.____, 1, wx.ALL, int(spacer))
            '''

            self.noFolderPopUp=wx.MessageDialog(self.panel, "You have not selected either a source or destination folder", "Error")

    #Class Methods  

        #Tries to populate dropdown with subdirs of \\pl2usext0008v0\Support. Excludes zip files, but they can be selected manually
        def dDownPopulate(self):
            support=r'\\pl2usext0008v0\Support'
            directories=[]
            #path=r'C:\Users\z004n7uy\Desktop\@'
            try:
                if os.path.isdir(support):
                    pathList=(os.listdir(support))
                    for items in pathList:
                        if "ZIPS" not in items.upper() and os.path.isdir(os.path.join(support,items)):
                            directories.append(os.path.join(support, items))
                    directories.reverse()
            except:
                directories=[r'\\pl2usext0008v0\Support\v2401', r'\\pl2usext0008v0\Support\v2306', r'\\pl2usext0008v0\Support\v2301', r'\\pl2usext0008v0\Support\v2022.2']
                #r'\\pl2usext0008v0\Support\v2022.1', r'\\pl2usext0008v0\Support\v2021.2',r'\\pl2usext0008v0\Support\v2021.1']      #Old dirs, no longer seem to be in use
            return directories

        #Returns the subdirs in the source dir      
        def populateListBox(self):
            lbEntries=[]
            try:
                for item in os.listdir(self.dirDDown.GetValue()):
                    if os.path.isdir(os.path.join(self.dirDDown.GetValue(), item)):
                        lbEntries.append(item)        
            except:
                errorPop=wx.MessageDialog(self.panel, "You do not currently have access to "+self.dirDDown.GetValue()+". Check your VPN connection", "Error")
                errorPop.ShowModal()
                self.dirDDown.SetValue('')
            return lbEntries

        #Populates listbox with valid subdirs            
        def updateListBox(self, evt):
            x=self.populateListBox()
            self.subdirSelect.Set(x)
            self.subdirChoices=x

        #Selects all options from the listbox
        def allSelect(self, evt):
            x=0
            for x in range (len(self.subdirChoices)):
                self.subdirSelect.SetSelection(x)
                x+=1             

        #Sets the listbox selection to nothing
        def clearSelect(self, evt):
            self.subdirSelect.SetSelection(wx.NOT_FOUND)

        #Selects the folders highlighted by the preview in the listbox
        def sendSelections(self, selects=[]):
            self.clearSelect('evt')
            for item in selects:
                self.subdirSelect.SetStringSelection(item)

        #Opens a folder dialog to select an alternate directory as the source dir. Calls to update the listbox to reflect this
        def altDir(self, event):
            FolderDialog=wx.DirDialog(self, "Choose directory to sync with", name="dirDialog")
            if FolderDialog.ShowModal()==wx.ID_CANCEL:
                return
            self.altDirName=FolderDialog.GetPath()
            self.dirDDown.SetValue(self.altDirName)
            self.updateListBox('evt')

        #Opens a folder dialog to select a destination dir
        def destDir(self, event):
            FolderDialog=wx.DirDialog(self, "Choose directory to sync to", name="dirDialog")
            if FolderDialog.ShowModal()==wx.ID_CANCEL:
                return
            self.destDirName=FolderDialog.GetPath()
            self.destDirText.SetValue(self.destDirName)

        #Returns the passed path truncated to the length of the source dir, used to create the save location path
        def dirUnplugger(self, dest, src=''):
            if src=='':
                src=self.dirDDown.GetValue()
            length=len(src)
            return dest[length:]

        #Opens the destination directory
        def openDestDir(self, evt):
            subprocess.Popen("explorer "+ self.destDirText.GetValue())

        #-Unfinished- 
        def scheduleTask(self, evt):
            return
            taskName=r'Support\Support Updater1'
            program=os.getcwd()+"\\SupportUpdater.pyw"
            username=os.getlogin()
            sys.stdout.write (username)
            subprocess.run(["SCHTASKS", r"/CREATE", r"/SC", 'MINUTE', r"/TN", taskName, r"/TR", program, r"/ST", "11:00", r"/RU", username])

        #Recursively calculates the size of a folder
        def getDirSize(self, dir):
            total=0
            try:
                for item in os.scandir(dir):
                    if os.path.isfile(item):
                        total+=item.stat().st_size
                    else:
                        total+=self.getDirSize(item.path)
            except:
                return 0
            return total

        #Returns true if the source directory was edited more recently than the destination or if it is larger
        def compareData(self, src, dst):
            srcStat=os.stat(src)
            dstStat=os.stat(dst)
            if (srcStat.st_mtime>dstStat.st_mtime):
                return True
            elif (self.getDirSize(src)>self.getDirSize(dst)):
                return True
            else:
                return False

        #Returns any files that exist only in root or differ across root and root2
        def getUniqueFiles(self, root, root2):
            uniqueFiles=[]
            try:
                self.workingPopUp.Pulse("Syncing files from \n"+str(os.path.basename(root2)))
                a=filecmp.dircmp(root, root2)
                self.workingPopUp.Pulse("Syncing files from \n"+str(os.path.basename(root2)))
                for item in a.left_only:
                    if not os.path.isdir(os.path.join(a.left, item)):
                        uniqueFiles.append(item)
                    self.workingPopUp.Pulse("Syncing files from \n"+str(os.path.basename(root2)))
            except FileNotFoundError:
                #sys.stdout.write ("root: "+root)
                #sys.stdout.write ("root2: "+root2)
                os.makedirs(root2)
                self.workingPopUp.Pulse("Syncing files from \n"+str(os.path.basename(root2)))
                a=filecmp.dircmp(root, root2)
                self.workingPopUp.Pulse("Syncing files from \n"+str(os.path.basename(root2)))
                for item in a.left_only:
                    if not os.path.isdir(os.path.join(a.left, item)):
                        uniqueFiles.append(item)
                    self.workingPopUp.Pulse("Syncing files from \n"+str(os.path.basename(root2)))
            uniqueFiles.extend(a.diff_files)
            return uniqueFiles

        #Copies given item from root to location
        def copyFiles(self, item, root, location):
            sys.stdout.write ("LOCATION : "+location+"\n")
            try:
                shutil.copy2(os.path.join(root, item), os.path.join(location, item))
                sys.stdout.write("Saved: "+ os.path.join(location, item)+"\n")
            except FileNotFoundError:
                os.mkdir(location)
                shutil.copy2(os.path.join(root, item), os.path.join(location, item))
                sys.stdout.write("Saved: "+ os.path.join(location, item)+"\n")
            except PermissionError:
                sys.stdout.write("No perms for "+os.path.join(location, item))

        #Returns a list of all subdirs in the destination that need to be updated
        def checkDatesInFolder(self):
            support=self.dirDDown.GetValue()
            dstDir=self.destDirText.GetValue()
            updateNeeded=[]
            selections=self.subdirSelect.GetStrings()
            for item in selections:
                if os.path.isdir(os.path.join(support, item)) or not os.path.exists(os.path.join(dstDir, item)):   
                    try:
                        if (self.compareData(os.path.join(support, item), os.path.join(dstDir, item))):
                            updateNeeded.append(item)
                    except:
                        sys.stdout.write ("ERROR COMPARING FILES. MAY NOT EXIST ")
                        updateNeeded.append(item)
            sys.stdout.write ("Update needed "+str(updateNeeded))
            return updateNeeded

        #Returns a list of syubdirs from the selection in the listbox that need updating
        def checkDatesInFolder2(self, list):
            support=self.dirDDown.GetValue()
            dstDir=self.destDirText.GetValue()
            updateNeeded=[]
            choices=self.subdirSelect.GetStrings()
            for item in list:
                if item ==0:
                    if self.compareData(support, dstDir):
                        updateNeeded.append(choices[item]) 
                else:
                    item=choices[item]
                    try:
                        if (self.compareData(os.path.join(support, item), os.path.join(dstDir, item))):
                            updateNeeded.append(item)
                    except:
                        sys.stdout.write ("ERROR IN COMPARING FILES. MAY NOT EXIST")
                        updateNeeded.append(item)
            sys.stdout.write ("Update needed "+str(updateNeeded))
            return updateNeeded

        #Checks for confirmation, clears the selected destination dir
        def destruct(self, evt):
            if self.isBlank():
                self.noFolderPopUp.ShowModal()
                return ''
            path=self.destDirText.GetValue()
            dlg=wx.TextEntryDialog(self.panel, "Type 'CONFIRM' to confirm deletion of \""+str(path)+"\"", "Confirm Deletion")
            result=dlg.ShowModal()
            sys.stdout.write (str(result))
            while dlg.GetValue()!="CONFIRM" and result!=wx.ID_CANCEL:
                error=wx.MessageDialog(self.panel, "The code you entered is incorrect", "Error")
                error.ShowModal()
                result=dlg.ShowModal()
            if result!=wx.ID_CANCEL:
                try:
                    shutil.rmtree(path, onerror=self.onError)
                    msg=wx.MessageDialog(self.panel, "Deletion in progress", "Deleting...")
                    msg.ShowModal()
                    os.mkdir(path)
                except:
                    errMsg=wx.MessageDialog(self.panel, str(traceback.print_exc()), "Error")
                    errMsg.ShowModal()
                    sys.stdout.write("Directory doesn't exist")

        #For shutil.rmtree call in destruct method. Updates read-only files with permissions needed to delete them
        def onError(self, func, path, exc_info):
            import stat
            if not os.access(path, os.W_OK):
                os.chmod(path, stat.S_IWUSR)
                func(path)
            else:
                raise


        #PReviews what files are out of date
        def previewSync(self, evt):
            if self.isBlank():
                self.noFolderPopUp.ShowModal()
                return ''
            if len(self.subdirSelect.GetSelections())==0:
                param=self.checkDatesInFolder()
            else:
                param=self.checkDatesInFolder2(self.subdirSelect.GetSelections())
            if len(param)==0 and len(self.subdirSelect.GetSelections())==0:
                emptyPopUp=wx.MessageDialog(self.panel, "All directories are up to date.")
                emptyPopUp.ShowModal()
            elif len(param)==0 and len(self.subdirSelect.GetSelections())!=0:
                emptyPopUp=wx.MessageDialog(self.panel, "All selected directories are up to date.")
                emptyPopUp.ShowModal()
            else:
                readable=''
                for item in param:
                    readable+=str(item)+"\n"
                previewMsg=wx.MessageDialog(self.panel, "The following directories/files from \""+str(self.dirDDown.GetValue())+"\" are out of date in \""+str(self.destDirText.GetValue())+"\":\n\n"+str(readable), style=wx.HELP)
                previewMsg.SetHelpLabel("Select above directories")
                if previewMsg.ShowModal()==5009:
                    self.sendSelections(param)

        #Syncs appropriate files
        def syncDirs(self, event, dir1=''):
            if self.isBlank():
                self.noFolderPopUp.ShowModal()
                return ''
            sys.stdout.write ("\n\n\n-----New-----\n\n")
            if len(self.subdirSelect.GetSelections())==0:
                updateNeededDirs=self.checkDatesInFolder()  
                sys.stdout.write ("No selections")
            else:
                updateNeededDirs=self.checkDatesInFolder2(self.subdirSelect.GetSelections())
                sys.stdout.write ("some selection")
            if len(updateNeededDirs)==0 and len(self.subdirSelect.GetSelections())==0:
                emptyPopUp=wx.MessageDialog(self.panel, "All directories are up to date.")
                emptyPopUp.ShowModal()
            elif len(updateNeededDirs)==0 and len(self.subdirSelect.GetSelections())!=0:
                emptyPopUp=wx.MessageDialog(self.panel, "All selected directories are up to date.")
                emptyPopUp.ShowModal()
            else:
                if dir1=='':
                    dir1=self.dirDDown.GetValue()
                dir2=self.destDirText.GetValue()
                self.dirZips=self.findZipsFolder()
                try:
                    permsCheck=(os.stat(dir1))
                    del permsCheck
                except PermissionError:
                    sys.stdout.write ("No Perms for "+ dir1+"\n")
                self.workingPopUp=wx.ProgressDialog(parent=self.panel, title="Syncing Files...", message="Syncing files from \n"+str(dir1), style=wx.PD_CAN_ABORT | wx.PD_SMOOTH | wx.PD_AUTO_HIDE, maximum=99999)
                self.workingPopUp.SetMinSize(wx.Size(5000,3000))
                self.workingPopUp.Center()
                self.workingPopUp.SetRange(1)
                tooBigDirs=[]
                for items in updateNeededDirs:
                    #sys.stdout.write (self.getDirSize(os.path.join(dir1,items)))
                    #if self.getDirSize(os.path.join(dir1,items))>1000000000 and self.checkFolderInZips(items):
                    if self.checkFolderInZips(items):
                        try:
                            self.copyZips(items)
                            #updateNeededDirs.remove(items)
                            continue
                        except:
                            pass
                    if self.workingPopUp.WasCancelled():
                        self.workingPopUp.Destroy()
                        break
                    sys.stdout.write ("                                 Root Dir---> "+os.path.join(dir1,items)+"\n")
                    
                    for root, _, _ in os.walk(os.path.join(dir1,items), followlinks=True):
                        if self.workingPopUp.WasCancelled():
                            self.workingPopUp.Destroy()
                            break
                        self.workingPopUp.Pulse("Syncing files from \n"+str(items))
                        uniqueFiles=[]
                        root2=dir2+self.dirUnplugger(root)
                        uniqueFiles=self.getUniqueFiles(root, root2)
                        location=dir2+self.dirUnplugger(root)
                        for item in uniqueFiles:
                            self.workingPopUp.Pulse("Syncing files from \n"+str(items))
                            self.copyFiles(item, root, location)
                            if self.workingPopUp.WasCancelled():
                                self.workingPopUp.Destroy()
                                exit
                                #break
                            self.workingPopUp.Pulse("Syncing files from \n"+str(items))
                        if self.workingPopUp.WasCancelled():
                            self.workingPopUp.Destroy()
                            exit
                            #break
                        else:
                            self.workingPopUp.Pulse("Syncing files from \n"+str(items))
                            shutil.copystat(root, location)
                self.workingPopUp.SetRange(1)
                self.workingPopUp.Update(1)
                self.workingPopUp.Destroy()          
                if self.workingPopUp.WasCancelled(): 
                    sys.stdout.write("\nCANCELLED") 
                else:
                    readable=''
                    for item in updateNeededDirs:
                        readable+=str(item)+"\n"
                    sys.stdout.write("\nDONE")
                    self.finishedPopUp=wx.MessageDialog(self.panel, "File sync complete. The following directories were updated:\n\n"+str(readable), "Complete")
                    self.finishedPopUp.ShowModal()
                #sys.exit()
                #sys.stdout.write (str(tooBigDirs))

        def findZipsFolder(self):
            support=self.dirDDown.GetValue()
            supportZip=support+" Zips"
            try:
                os.stat(supportZip)
            except:
                pass
            try:
                supportZip=support+" ZIPS"
                os.stat(supportZip)
            except:
                return ''
            #sys.stdout.write ("\nFound zips folder: "+supportZip)
            return supportZip
        
        def checkFolderInZips(self, subdir):
            #sys.stdout.write (os.listdir(self.findZipsFolder()))
            try:
                return subdir+".zip" in os.listdir(self.dirZips)
            except:
                return False

        def copyZips(self, subdir):
            self.workingPopUp.Pulse("Finding zip folder")
            dest=self.destDirText.GetValue()
            tempDir=os.path.join(dest,"ZIPtemp")
            try:
                supportZip=self.dirZips
            except:
                supportZip=self.findZipsFolder()
                sys.stdout.write("self.dirZips not found")
            self.workingPopUp.Pulse("Making temporary zip folder on local machine\n ")
            if not os.path.exists(tempDir):
                os.mkdir(tempDir)
                f=open(os.path.join(tempDir,"WhatIsThisFolder.txt"), 'w')
                f.write ("This folder is a temporary folder made by the Support Updater tool to facilitate the use of ZIP files to aid in overhead times for the tool.\nThis folder should be deleted after a successful run of the tool, but a failure in the tool ending properly may result in this file persisting. \nFeel free to delete manually. No information of yours is stored here.")
                f.close()
                sys.stdout.write ("Made " + tempDir)
            #zips=[]
            self.workingPopUp.Pulse("Copying zip to local\n"+subdir+".zip")
            try:
                shutil.copy2(os.path.join(supportZip,subdir)+".zip", tempDir)
            except:
                sys.stdout.write ("Couldnt find zip file for "+os.path.join(supportZip,subdir)+".zip\n\n")
                raise Exception
            self.workingPopUp.Pulse("Copying zip to local\n"+subdir+".zip")
            #sys.stdout.write ("Copied "+os.path.join(supportZip,subdir)+".zip to "+tempDir)
            #zips.append(os.path.join(supportZip,item)+".zip")
            self.workingPopUp.Pulse("Unzipping folder\n"+subdir+".zip")
            zipObj=ZipFile(os.path.join(supportZip,subdir)+".zip", 'r')
            zipObj.extractall(tempDir)
            self.workingPopUp.Pulse("Unzipping folder\n"+subdir+".zip")
            #sys.stdout.write ("Extracted " +os.path.join(supportZip,subdir)+".zip to "+tempDir)
            uniqueFiles=self.getUniqueFiles(os.path.join(tempDir,subdir), os.path.join(dest,subdir))
            sys.stdout.write ("                                 Root Dir---> "+os.path.join(tempDir,subdir))
            dir2=os.path.join(dest,subdir)
            self.workingPopUp.Pulse("Syncing files from\n"+subdir)
            for root, _, _ in os.walk(os.path.join(tempDir,subdir), followlinks=True):
                if self.workingPopUp.WasCancelled():
                    self.workingPopUp.Destroy()
                    break
                #uniqueFiles=[]
                #dir2=dest
                root2=dir2+self.dirUnplugger(root, os.path.join(tempDir,subdir))
                #if os.path.basename(root)!=item:
                    #location=os.path.join(root2, os.path.basename(root))
                location=root2

                uniqueFiles=self.getUniqueFiles(root, root2)
                location=root2
                for file in uniqueFiles:
                    self.copyFiles(file, root, location)
                    if self.workingPopUp.WasCancelled():
                        self.workingPopUp.Destroy()
                        exit
                        #break
                    self.workingPopUp.Pulse("Syncing files from \n"+str(subdir))
                if self.workingPopUp.WasCancelled():
                    self.workingPopUp.Destroy()
                    exit
                    #break
                else:
                    self.workingPopUp.Pulse("Syncing files from \n"+str(subdir))
                    shutil.copystat(root, location)
            shutil.rmtree(tempDir, self.onError)    
             
        #Checks if source and destination fields are empty
        def isBlank(self):
            return (self.destDirText.GetValue()=='' or self.dirDDown.GetValue()=='')


    #Main method
    if __name__=="__main__":
        vsc = win32gui.GetForegroundWindow()
        if (sys.argv[0][-4:]==".exe"):
            win32gui.ShowWindow(vsc, win32con.SW_HIDE)
        app = wx.App()
        top = compareFrame("Support Updater Tool")
        try:
            icon=wx.Icon("nxfemap_app.ico", wx.BITMAP_TYPE_ICO)
            top.SetIcon(icon)
        except:
            sys.stdout.write("Icon not found. Place .ico file in same folder as this .exe to enable")
        top.Show()
        top.Center()
        app.MainLoop()
        #show = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(vsc, win32con.SW_SHOW)



else:
    a=ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    #sys.stdout.write("Administrator privileges not detected. Restart with admin privileges\n")
