import tkinter as tk
import tkinter.filedialog as fileDialog
import tkinter.messagebox as messageBox
from tkinter import *
from tkinter.filedialog import askopenfilename   
import shlex, subprocess, shutil, os, platform, posixpath, shlex, time , _thread, sys
from subprocess import Popen, PIPE, STDOUT

class Application():

    # Initilizes global variable and stores user selected file path as the source
    def selectHotfolder():
        global src
        src = fileDialog.askdirectory()
        root.destroy()
        

    def startScript():
        global monitoring
        global stopMonitor
        global queueBusy
        global folderToMonitor
        global WatchFolder
        global root
        #os.system('cls' if os.name == 'nt' else 'clear')
        os.system('mode con: cols=220 lines=90')
        monitoring = False
        stopMonitor = False
        queueBusy = False
        copyThreadDone = False
        WatchFolder = False
        folderToMonitor = ""

        # Get file directory of script
        fileDir = os.path.dirname(os.path.realpath(__file__))
            
        # Join root path with TAG folder
        infile = open(os.path.join(fileDir,"aboutMe.txt"),"r")

        for eachLine in infile:
            if eachLine != "\n":
                print (eachLine,end="\r")

        userInput = input("\nHit ENTER to select HotFolder...")

        fname = "unassigned"
        
        if __name__ == '__main__':

            root = Tk()
            Button(root, text='Select HotFolder', command = Application.selectHotfolder).pack(fill=X)
            mainloop()
            
            print("\nHot folder selected:\n" + "'"    + src + "'")

        userInput = input("\nHit ENTER to START monitoring folder and begin print process.\n")
        Application.printStart()
       
    def progressReport(message,delay):
        try:
            for x in range (0,4):  
                b = str(message) + "." * x
                print ('\r'+b) 
                time.sleep(delay)
                os.system('CLS')
        except:
            monitoring = False

    def getFile(fromQueue):
        
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

                        endLoop = False
                        
                        while endLoop == False:

                            currentFile = allFiles[e]                         
                            if currentFile[2] == fromQueue:
                                #DEBUG
                                #print(currentFile)
                                fileToReturn = currentFile
                                endLoop = True
                                gotFile = True

                            e += 1

                            if e == i:
                                endLoop = True
                            
                        if gotFile == False:
                            fileToReturn = ""
                            gotFile = True
            else:
                statusReport = "Processing files..."
                Application.printFilelist(2)
                                              
        if gotFile == True:
            return fileToReturn
        else:
            return ""
        #DEBUG
        #print ("Got File")


    def addSpacing(text, slotType):

        separator = " | "

        if slotType == "slot":
            columnLenght = 4
        elif slotType == "name":
            columnLenght = 75
        elif slotType == "status":
            columnLenght = 12

            
        spacesToAdd = columnLenght - len(text)

        if spacesToAdd > 0:
            newText = text + " " * spacesToAdd + separator

        elif spacesToAdd <= 0:
            newText = text + separator

        return newText

    def printFilelist(delay=5):

        totalFiles = len(allFiles)
        os.system('CLS')
        print("Hotfolder to Watch: " + src + "\n\n")
        for i in range(0,totalFiles):

            file = allFiles[i]
            slot = Application.addSpacing(str(file[0]),"slot")
            name = Application.addSpacing(str(file[1]),"name")
            status = Application.addSpacing(str(file[2]),"status")
            try:
                print(slot + name + status)
            except:
                print("\n\n")
                exit()
                
        print("\n    Status:\n         " + statusReport)
        time.sleep(delay)

    #
    # Starts the file copy process
    #
    def printStart():

        # To stop monitor once process completes
        global stopMonitor
        
        # used to store and report copiedFiled
        global copiedFiles
        
        global allFiles

        global statusReport
        
        #Initiate stopMonitor and Monotoring boolean to stop the monitor service, not entirely necessary. Will deprecate soon.
        stopMonitor = False

        # Start Monitoring the Hotfolder using a subprocess
        Application.startMonitoring()

        statusReport = "Finalizing start..."
        
        
        #Wait for the monoitor to start and complete file checks.
        while monitoring != True:

            Application.progressReport("Analyzing folder and starting file monitor ",.5)

        #Initiate file variables
        file_name = ""
        theFile = []

        while monitoring == True:
            
            #DEBUG: Show user file list
            #print(allFiles)

            if file_name != "":
            
                file_name_src = os.path.join(src, file_name)

                statusReport = "Print from: " + file_name_src
                Application.printFilelist(2)

                printedSuccesfully = False

                # Begin Print Procedure
                if os.path.exists(file_name_src) == True:
                
                    # If the path is a file
                    if (os.path.isfile(file_name_src)):

                        statusReport = "Printing..."
                        try:

                            # Use GSPrint Ghostscript to handle PDF printing
                            #DEBUG: Comment this section out to prevent excessive printing when testing
                            gsprint = r"C:\Program Files\Ghostgum\gsview\gsprint.exe"
                            cmd = '"{}" -noquery -landscape "{}"'.format(gsprint, file_name_src)
                            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            stdout, stderr = proc.communicate()
                            exit_code = proc.wait()
                            #END DEBUG
                            
                            statusReport = "Printed!"
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
                
                theFile = Application.getFile("Available")

                #self.progressReport("Waiting for file",.5)
                if len(allFiles) > 0:
                        statusReport = "Waiting for file..."
                        Application.printFilelist()

   
                if theFile != "":
                
                    file_name = theFile[1]
                    statusReport = "Got file " + str(file_name) + " from queue!"
                    Application.printFilelist()
                    
                
        print ("Print Process Complete\n")

            

    # Quits the script
    def quitScript():
        print("Quit")

    def startMonitoring():
        
        # Create new thread
        try:
            
            _thread.start_new_thread(Application.fileMonitor, (src, 2, ) )
            
        except:
            
            print ("Error: unable to start thread")

    def print_time(threadName, delay):
        count = 0
        while count < 5:
            time.sleep(delay)
            count += 1
            print ("%s: %s" % ( threadName, time.ctime(time.time()) ))
    
    def files_to_timestamp(path):
        files = [os.path.join(path, f) for f in os.listdir(path)]
        return dict ([(f, os.path.getmtime(f)) for f in files])

    def get_size(start_path = '.'):
            
            total_size = 0
            for dirpath, dirnames, filenames in os.walk(start_path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    total_size += os.path.getsize(fp)

    
    def incrementType(stream):
        
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

    def isFileInUse(start_path = '.', printResults= True,stream= False, pollTime = 5):
    
        t_end = time.time() + pollTime
        
        sizeOfFile = []
        if stream ==True:

            print(start_path)
        
        while time.time() < t_end:
            
            size = Application.get_size(start_path)
            
            sizeOfFile.append(size)
            
            if stream ==True:
                
                print(size)
            
            if printResults==True:
                
                Application.progressReport("\tPolling",.5)
                
            else:
                
                time.sleep(.5)
    
        dirIncrement = Application.incrementType(sizeOfFile)
                
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

    def checkAllFiles(src,arrayOfFiles):
        
        i = len(arrayOfFiles)
        e = 0
        
        filesToReturn = []

        if i != 0:
        
            while e < i:
                
                if arrayOfFiles[e] != ".DS_Store":

                    fileName = arrayOfFiles[e]
                    filePath = os.path.join(src,arrayOfFiles[e])
                    
                    if Application.isFileInUse(filePath, printResults= False, pollTime = 6):
                    
                        fileInUse = "Busy"
                        
                    else:
                        fileInUse = "Available"
                    
                    
                    filesToReturn.append((e,fileName,fileInUse))
                
                e += 1
            
                print("File Added")
        
        return filesToReturn


    def addToQueue(itemsToAdd):
        
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
                    fileStat = Application.isFileInUse(start_path=filePath, printResults= False,stream=False, pollTime = 5)
                    
                    slotNumber = len(allFiles)
                
                    if fileStat:
                    
                        fileInUse = "Busy"
                        
                    else:
                        
                        if fileSize > 0:
                        
                            fileInUse = "Available"
                        else:
                            fileInUse = "Busy"
                        
                    allFiles.append((slotNumber,fileName,fileInUse))
                    
                    statusReport = "Added: " + fileName
                    Application.printFilelist(2)
            
            e += 1

    def removeFromQueue(itemsToRemove):

        i = 0
        
        while i < len(itemsToRemove):
            
            fileName = os.path.basename(itemsToRemove[i])
            
            Application.updateItem(fileName, fileStat="Removed")
                    
            i += 1

    def modifiedInQueue(itemsModified):
    
        i = 0
        
        while i < len(itemsModified):
            
            fileName = os.path.basename(itemsModified[i])
            
            Application.updateItem(fileName, fileStat="Modified")
                    
            i += 1

    def updateItem(fileName,fileSlot = 0,fileStat=""):
        
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

            statusReport = fileStat + ":" + fileName
            Application.printFilelist(2)
            
        #else:
            
            #print ("\tFailed to update status: " + fileName)


    def fileMonitor(srcToWatch, delay):
        
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
        #print(availableFiles)

        before = Application.files_to_timestamp(path_to_watch)
        
        print ("Watch Start:" + path_to_watch + "\n")
        
        while stopMonitor != True:
            
            if filesChecked ==False:

                #DEBUG
                #print("Checking all files")
                
                allFiles = Application.checkAllFiles(srcToWatch, availableFiles)
                
                #DEBUG
                #print("All files checked")
                
                filesChecked = True
                
                monitoring = True

                #DEBUG: Print all queue files
                #print(allFiles)
                
                time.sleep(delay*4)
                
            after = Application.files_to_timestamp(path_to_watch)
            added = [f for f in after.keys() if not f in before.keys()]
            removed = [f for f in before.keys() if not f in after.keys()]
            modified = []
        
            for f in before.keys():
                if not f in removed:
                    if os.path.getmtime(f) != before.get(f):
                        modified.append(f)
                        
            queueBusy = True
            if added:
                Application.addToQueue(added)
            if removed:
                Application.removeFromQueue(removed)
            if modified: 
                Application.modifiedInQueue(modified)
            queueBusy = False

            
            time.sleep (delay)
            
            before = after

            #DEBUG: Print all queue files
            #print("Monitoring")
        
        print ("Watch End: " + path_to_watch + "\n")
        
        monitoring = False
#try:
Application.startScript()
#except:
#    print("\n\nFile monitor and Script have been stopped")
