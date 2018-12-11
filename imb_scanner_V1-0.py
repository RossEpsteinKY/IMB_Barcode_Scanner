import time
import os.path
import glob
import zipfile
import os
import win32print
import fnmatch
import datetime

now = datetime.datetime.now()


def StartupProgram():
    os.system('cls')
    print("Welcome to IMB Scanner")
    time.sleep(3)
    os.system('cls')
    DriveSelect()


def DriveSelect():
    global drive
    drive = str(input("Please select the desired drive letter? (ie: W, X, Y, Z) >  "))
    drive = drive + str(":/")
    os.system('cls')
    print("You have selected drive: " + drive)
    time.sleep(2)
    os.system('cls')
    UnzipFiles()

def UnzipFiles():
    global new_dr
    new_dr = str(drive) + "ziptest/todays_pdrs"
    

    #def UnzipComplete():
    #    os.system('cls')
    #    print("Successfully unzipped all files.")
    #    time.sleep(3)
    #    os.system('cls')
    #    BarcodeScanner()



    def archive1():
        global number_files
        global filecounter
        new_dr = str(drive) + "MAIL DAT EXPORTS/READY FOR MAIL DAT SCAN CHECK"  #insert path to zips
        filelist = os.listdir(new_dr) 
        number_files = len(filelist)
        number_files = int(number_files)
        
        filecounter = 0
        rootPath= (str(new_dr))
        pattern = '*.zip'
        while int(filecounter) < int(number_files):
            for root, dirs, files in os.walk(rootPath):
                for filename in fnmatch.filter(files,pattern):
                    with zipfile.ZipFile(os.path.join(root, filename)) as zf:
                        extractor(zf)
                        print("Extracting " + str(filename))
                        filecounter = filecounter + 1
        os.system('cls')
        print(str(number_files) + " *.pdr files were successfully unzipped into " + str(new_dr))
        time.sleep(3)
        os.system('cls')
        print("Loading barcode scanner...")
        time.sleep(2)
        os.system('cls')
        BarcodeScanner()

    def extractor(zip_file):
        global extensions
        new_dr_unzip = r'.'
        extensions = ('.pdr')
        [zip_file.extract(file,new_dr_unzip) for file in zip_file.namelist() if file.endswith(extensions)]
    
    if __name__ == '__main__':
        archive1()
  
def BarcodeScanner():
    global scan_drive

    scan_drive = str(drive) + "."
    print("Barcode scanner loaded")
    print("Scanning in folder: " + scan_drive)
    time.sleep(3)
    os.system('cls')
    ScannerStartpoint()



def ScannerStartpoint():
    global searchstring
    global safeToShip
    global b
    b = 0
    try:
        b = int(input("What is the barcode?  >"))
    except:
        print("You must input only intergers")
        time.sleep(2)   
        os.system('cls') 
        ScannerStartpoint()
    

    
    searchstring = b
    for fname in os.listdir('.'):
        if fname.endswith(extensions):
            #print(str(len(fname)))    
            if os.path.isfile(fname):    
                with open(fname) as f:   
                    for line in f:       
                        if str(b) in line:
                            global success_alert
                            success_alert = str(fname)
                            global foundFileIn
                            foundFileIn = str(success_alert) 
                            safeToShip = True
                            CheckIfSafe()
                        else:
                            safeToShip = False
                        
    if safeToShip == False:
        CheckIfSafe()
    elif CheckIfSafe == True:
        CheckIfSafe()
                         

def CheckIfSafe():
    if safeToShip == True:
        print("BARCODE FOUND IN FILE " + str(foundFileIn) + "!")
        print("IT IS SAFE TO SHIP BARCODE #" + str(searchstring))
        print("Preparing to print success letter...")
        time.sleep(1)   
        os.system('cls')
        #printout_name = str(foundFileIn) + ".txt"
        printout_success = "successfile.txt"

        file = open(printout_success,"w") 
        found_date = now.strftime("%Y-%m-%d %H:%M")
        file.write("SUCCESSFULLY FOUND BARCODE #" + str(searchstring)) 
        file.write("\nSUCCESSFULLY FOUND BARCODE IN JOB FILE " + str(foundFileIn)) 
        file.write("\nSuccessfully found at " + str(found_date))
        file.write("\nAPPROVED TO SHIP") 
        file.close()
        os.system('notepad /p successfile.txt')   
        ScannerStartpoint()

    elif safeToShip == False:
        print("UNDOCUMENTED MAIL! BARCODE NOT FOUND!")
        print("DO NOT SHIP BARCODE #" + str(searchstring))
        time.sleep(3)   
        os.system('cls')  
        printout_failure = "failurefile.txt"

        file = open(printout_failure,"w") 
        found_date = now.strftime("%Y-%m-%d %H:%M")
        file.write("\n*************************************************")
        file.write("\nUNDOCUMENTED MAIL! BARCODE NOT FOUND!")
        file.write("\nBARCODE #" + str(searchstring) + " IS NOT IN SYSTEM") 
        file.write("\nDO NOT SHIP THIS MAIL")
        file.write("\nFAILED SCAN OCCURED AT " + str(found_date))
        file.write("\n*************************************************")
        file.close()
        os.system('notepad /p failurefile.txt')
        ScannerStartpoint()
    else:
        print("An error has occured")

    print("Barcode scanner loaded")
    print("Scanning in folder: " + scan_drive)
    time.sleep(3)
    os.system('cls')
    ScannerStartpoint()                          

StartupProgram()
