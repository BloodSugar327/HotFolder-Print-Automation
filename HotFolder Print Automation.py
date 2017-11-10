import tkinter as tk
import tkinter.filedialog as fileDialog
import tkinter.messagebox as messageBox
import shlex, subprocess, shutil, os, platform, posixpath, shlex, time , _thread, sys
from subprocess import Popen, PIPE, STDOUT

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        global monitoring
        global monitorStatus
        global stopMonitor
        global queueBusy
        global folderToMonitor
        global WatchFolder
        os.system('cls' if os.name == 'nt' else 'clear')
        self.pack()
        self.createWidgets()     
        monitoring = False
        stopMonitor = False
        queueBusy = False
        copyThreadDone = False
        WatchFolder = False
        folderToMonitor = ""
        
        print("Script Started Succesfully...\n\n* Select Folders to copy in GUI *")
    
    # Create 4 form buttons
    def createWidgets(self):
        # Button to select source folder
        self.selectSrcButton = tk.Button(self, text="Select HotFolder", command=self.selectSource_folder)
        self.selectSrcButton.pack(side="left")
        
        # Button to start file print
        global startButtonText
        startButtonText = tk.StringVar()
        self.Start = tk.Button(self, textvariable=startButtonText, fg="green", command=self.printStart)
        self.Start.pack(side="left")
        startButtonText.set("Start")
        
        
    def progressReport(self,message,delay):
        try:
            for x in range (0,4):  
                b = str(message) + "." * x
                print ('\r'+b) 
                time.sleep(delay)
                os.system('CLS')
        except:
            monitoring = False

    # Initilizes global variable and stores user selected file path as the source
    def selectSource_folder(self):
        global src
        src = fileDialog.askdirectory()
        print("\nHot folder selected:\n" + "'"    + src + "'")



    def getFile(self, fromQueue):
        
        global allFiles
        global queueBusy
        
        gotFile = False
        
        fileToReturn = ""

        #DEBUG:
        #print (allFiles)
        
        while gotFile == False:
            
            if queueBusy == False:
                
                i = len(allFiles)

                if i == 0:

                    fileToReturn = ""
                    gotFile = True

                else:
                    if i > 0 :
                        
                        e = 0
                        
                        while e < i:        
                            currentFile = allFiles[e]                         
                            if currentFile[2] == fromQueue:
                                #DEBUG
                                #print(currentFile)
                                fileToReturn = currentFile
                                gotFile = True
                                
                            e +=1
                            
                        if gotFile == False:
                            fileToReturn = ""
                            gotFile = True
            else:
                self.progressReport("Processing files",.5)
                                              
        if gotFile == True:
            return fileToReturn
        else:
            return ""
        #DEBUG
        #print ("Got File")


    def addSpacing(self,text, slotType):

        separator = " | "

        if slotType == "slot":
            columnLenght = 4
        elif slotType == "name":
            columnLenght = 50
        elif slotType == "status":
            columnLenght = 12

            
        spacesToAdd = columnLenght - len(text)

        if spacesToAdd > 0:
            newText = text + " " * spacesToAdd + separator

        elif spacesToAdd <= 0:
            newText = text + separator

        return newText

    def printFilelist(self):

        totalFiles = len(allFiles)

        for i in range(0,totalFiles):

            file = allFiles[i]
            slot = self.addSpacing(str(file[0]),"slot")
            name = self.addSpacing(str(file[1]),"name")
            status = self.addSpacing(str(file[2]),"status")
            try:
                print(slot + name + status)
            except:
                print("\n\n")
                exit()


    #
    # Starts the file copy process
    #
    def printStart(self):


        os.system('mode con: cols=200 lines=40')

        # To stop monitor once process completes
        global stopMonitor
        
        # used to store and report copiedFiled
        global copiedFiles
        
        global allFiles
        
        #Initiate stopMonitor and Monotoring boolean to stop the monitor service, not entirely necessary. Will deprecate soon.
        stopMonitor = False

        # Start Monitoring the Hotfolder using a subprocess
        self.startMonitoring()

        #Wait for the monoitor to start and complete file checks.
        while monitoring != True:

            self.progressReport("Waiting for file monitor to start",.5)

        #Initiate file variables
        file_name = ""
        theFile = []

        while monitoring == True:
            
            #DEBUG: Show user file list
            #print(allFiles)

            if file_name != "":
            
                file_name_src = os.path.join(src, file_name)

                print("\tPrint from: " + file_name_src)

                printedSuccesfully = False

                # Begin Print Procedure
                if os.path.exists(file_name_src) == True:
                
                    # If the path is a file
                    if (os.path.isfile(file_name_src)):

                        print("\nPrinting...")
                        try:

                            # Use GSPrint Ghostscript to handle PDF printing
                            #DEBUG: Comment this section out to prevent excessive printing when testing
                            gsprint = r"C:\Program Files\Ghostgum\gsview\gsprint.exe"
                            cmd = '"{}" -noquery -landscape "{}"'.format(gsprint, file_name_src)
                            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            stdout, stderr = proc.communicate()
                            exit_code = proc.wait()
                            #END DEBUG
                            
                            print("\nPrinted!")
                            theFile = (theFile[0],theFile[1],"Printed")
                            fileIndex = theFile[0]
                            allFiles[fileIndex] = theFile
                            theFile = ""
                            file_name = ""
                            
                        except:
                            
                            print("Error!\n")
                
            while file_name == "":

                #DEBUG:   
                #print(theFile)
                
                theFile = self.getFile("Available")

                #self.progressReport("Waiting for file",.5)
                if len(allFiles) > 0:
                        os.system('CLS')
                        self.printFilelist()
                        time.sleep(5)

   
                if theFile != "":
                
                    file_name = theFile[1]
                
                    print("\nGot file " + str(file_name) + " from queue!")
                 
        stopMonitor = True
        
        time.sleep(5)
                
        print ("Print Process Complete\n")

            

    # Quits the script
    def quitScript(self):
        print("Quit")

    def startMonitoring(self):
        
        # Create new thread
        try:
            
            _thread.start_new_thread(self.fileMonitor, (src, 2, ) )
            
        except:
            
            print ("Error: unable to start thread")

    def print_time(self, threadName, delay):
        count = 0
        while count < 5:
            time.sleep(delay)
            count += 1
            print ("%s: %s" % ( threadName, time.ctime(time.time()) ))
    
    def files_to_timestamp(self,path):
        files = [os.path.join(path, f) for f in os.listdir(path)]
        return dict ([(f, os.path.getmtime(f)) for f in files])

    def get_size(self,start_path = '.'):
            
            total_size = 0
            for dirpath, dirnames, filenames in os.walk(start_path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    total_size += os.path.getsize(fp)

    
    def incrementType(self,stream):
        
        i = len(stream)
        e = 0
        
        increasing = 0
        fixed = 0
        decreasing = 0
        
        topNumber = stream[e]
            
        e += 1
        
        while e < i:
            nextNumber = stream[e]
            
            if topNumber == nextNumber:
                fixed += 1
            elif topNumber > nextNumber:
                decreasing += 1
            elif topNumber < nextNumber:
                increasing += 1
                
            topNumber = nextNumber
                
            e += 1
                    
        increasingPercent = increasing/(i-1) * 100
        fixedPercent = fixed/(i-1) * 100
        decreasingPercent = decreasing/(i-1) * 100
                
        results = [
                ("Increasing", increasingPercent, increasing),
                ("Fixed", fixedPercent, fixed),
                ("Decreasing", decreasingPercent, decreasing),
        ]
                
        resultsSorted = sorted(results,key=lambda type: type[2],reverse=True)
        
        topResult = resultsSorted[0]
        
        return topResult

    def isFileInUse(self,start_path = '.', printResults= True,stream= False, pollTime = 5):
    
        t_end = time.time() + pollTime
        
        sizeOfFile = []
        if stream ==True:

            print(start_path)
        
        while time.time() < t_end:
            
            size = self.get_size(start_path)
            
            sizeOfFile.append(size)
            
            if stream ==True:
                
                print(size)
            
            if printResults==True:
                
                self.progressReport("\tPolling",.5)
            else:
                
                time.sleep(.5)
    
        dirIncrement = self.incrementType(sizeOfFile)
                
        if stream==True:

            print(dirIncrement)
        
        if dirIncrement[0] == "Fixed":
            
            if printResults==True:
                
                print("\n\tDirectory not in use")
                
            return False
            
        else:
            
            if printResults==True:
                
                print("\n\tDirectory is in use")
                
            return True

    def checkAllFiles(self,src,arrayOfFiles):
        
        i = len(arrayOfFiles)
        e = 0
        
        filesToReturn = []

        if i != 0:
        
            while e < i:
                
                if arrayOfFiles[e] != ".DS_Store":

                    fileName = arrayOfFiles[e]
                    filePath = os.path.join(src,arrayOfFiles[e])
                    
                    if self.isFileInUse(filePath, printResults= False, pollTime = 6):
                    
                        fileInUse = "Busy"
                        
                    else:
                        fileInUse = "Available"
                    
                    
                    filesToReturn.append((e,fileName,fileInUse))
                
                e += 1
            
                print("File Added")
        
        return filesToReturn


    def addToQueue(self,itemsToAdd):
        
        global allFiles
    
        i = len(itemsToAdd)
        e = 0
        
        foundIt = False
        
        while e < i:
            
            fileName = os.path.basename(itemsToAdd[e])
            
            if fileName != ".DS_Store":
            
                f = len(allFiles)
                g = 0
                
                while g < f:
                    
                    currentFile = allFiles[g]
                    
                    if currentFile[1] == itemsToAdd[e]:
                        
                        fileIndex = currentFile[0]
                        
                        newFile = (currentFile[0],currentFile[1],"Available")
                        
                        allFiles[fileIndex] = newFile
                        
                        foundIt = True
                        
                        
                    g += 1
                        
                if foundIt != True:
                                        
                    fileName = os.path.basename(itemsToAdd[e])
                    
                    filePath = os.path.join(src,itemsToAdd[e])
                    fileSize = os.path.getsize(filePath)
                    fileStat = self.isFileInUse(start_path=filePath, printResults= False,stream=False, pollTime = 5)
                    
                    slotNumber = len(allFiles)
                
                    if fileStat:
                    
                        fileInUse = "Busy"
                        
                    else:
                        
                        if fileSize > 0:
                        
                            fileInUse = "Available"
                        else:
                            fileInUse = "Busy"
                        
                    allFiles.append((slotNumber,fileName,fileInUse))
                    
                    print ("\t Added: " + fileName)
                    
            
            e += 1

    def removeFromQueue(self,itemsToRemove):

        i = 0
        
        while i < len(itemsToRemove):
            
            fileName = os.path.basename(itemsToRemove[i])
            
            self.updateItem(fileName, fileStat="Removed")
                    
            i += 1

    def modifiedInQueue(self,itemsModified):
    
        i = 0
        
        while i < len(itemsModified):
            
            fileName = os.path.basename(itemsModified[i])
            
            self.updateItem(fileName, fileStat="Modified")
                    
            i += 1

    def updateItem(self,fileName,fileSlot = 0,fileStat=""):
        
        global allFiles
        
        i = len(allFiles)
        e = 0
        
        found = False
        
        while e < i:
            
            currentItem = allFiles[e]
            
            if currentItem[1] == fileName:
                #Collect initial data, if empty in function
                
                if fileSlot == 0:
                    fileSlot = currentItem[0]
                if fileStat == "":
                    fileStat = currentItem[2]
                    
                found = True
            
            e += 1
        
        if found:
                
            newItem = (fileSlot,fileName,fileStat)
            
            allFiles[fileSlot] = newItem
                
            #print ("\t" + fileStat + ":" + fileName)
            
        #else:
            
            #print ("\tFailed to update status: " + fileName)


    def fileMonitor(self,srcToWatch, delay):
        
        global monitoring
        global allFiles
        global filesChecked
        global queueBusy
        
        filesChecked = False

        #DEBUG
        #print("Monitor Started")
        
        path_to_watch = os.path.abspath(srcToWatch)
    
        availableFiles = os.listdir(srcToWatch)
        
        #DEBUG: Used to view available files in folder so far
        print(availableFiles)

        before = self.files_to_timestamp(path_to_watch)
        
        print ("Watch Start:" + path_to_watch + "\n")
        
        while stopMonitor != True:
            
            if filesChecked ==False:

                #DEBUG
                #print("Checking all files")
                
                allFiles = self.checkAllFiles(srcToWatch, availableFiles)
                
                #DEBUG
                #print("All files checked")
                
                filesChecked = True
                
                monitoring = True

                #DEBUG: Print all queue files
                #print(allFiles)
                
                time.sleep(delay*4)
                
            after = self.files_to_timestamp(path_to_watch)
            added = [f for f in after.keys() if not f in before.keys()]
            removed = [f for f in before.keys() if not f in after.keys()]
            modified = []
        
            for f in before.keys():
                if not f in removed:
                    if os.path.getmtime(f) != before.get(f):
                        modified.append(f)
                        
            queueBusy = True
            if added:
                self.addToQueue(added)
            if removed:
                self.removeFromQueue(removed)
            if modified: 
                self.modifiedInQueue(modified)
            queueBusy = False

            
            time.sleep (delay)
            
            before = after

            #DEBUG: Print all queue files
            #print("Monitoring")
        
        print ("Watch End: " + path_to_watch + "\n")
        
        monitoring = False

try:
    root = tk.Tk()
    app = Application(master=root)
    app.master.title("PRIMARY Print Automation")
    app.mainloop()
except:

    print("Exit")
