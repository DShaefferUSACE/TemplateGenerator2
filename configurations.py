######################################
##  ------------------------------- ##
##     SAW Template Generator       ##
##  ------------------------------- ##
##     Written by: David Shaeffer   ##
##  ------------------------------- ##
## Last Edited on: 28-February-2019 ##
##  ------------------------------- ##
######################################

import os
from datetime import datetime, timedelta
from shutil import copy, rmtree
# import tkMessageBox
# import tkinter.messagebox
import getpass
import json
try:
    import Tkinter as tk
    from Tkinter import *
except ImportError:
    import tkinter as tk
    from tkinter import *

import subprocess
import sys

#parse json
def getPath(value):
    # get path to current working directory
    cwd = os.getcwd()
    dir_path = os.path.abspath(cwd + r'\Configuration_Files\services.json')
    with open(dir_path) as json_data:
        d = json.load(json_data)
    return d[value]

# Parse json and return user value
def getSettings(value):
    # get path to current working directory
    cwd = os.getcwd()
    dir_path = os.path.abspath(cwd + r'\Configuration_Files\usersettings.json')
    with open(dir_path) as json_data:
        d = json.load(json_data)
    return d[value]

## This removes any old temporary MEI folders
try:
    # get path to current working directory
    cwd = os.getcwd()
    all_subdirs = [d for d in os.listdir(cwd) if os.path.isdir(d)]
    # print all_subdirs
    latest_subdir = max(all_subdirs, key=os.path.getctime)
    # print latest_subdir
    ## Search for MEI folder that we not previously deleted 
    for item in os.listdir(cwd):
        if '_MEI' in item:
            if item != latest_subdir:
                rmtree(cwd + "/" + item)
except Exception as e:
    print (e)

# Step 1. check server for a more recent .exe file:
# IF a new file exists then copy it to the users configurations directory. Once it is done copying then open the new .exe and close the old one.
# # ELSE check for old versions of the .exe and delete then open the newest .exe
#Updates template generator .exe from server
try:
    print ("Searching for new version of the Template Generator..."),
    serverpath = getPath('TGfolder')
    # get path to current working directory
    cwd = os.getcwd()
    # # Search through the users working directory to find the .exe
    for item in os.listdir(cwd):
        if item.endswith(".exe"):
            localtime = os.path.getmtime(cwd + "/" + item)
    for item in os.listdir(serverpath):
        if item.endswith(".exe"):
            servertime = os.path.getmtime(serverpath + "/" + item)
    # print "Server Time: ",
    # print servertime
    # print "Local Time: ",
    # print localtime
    if servertime > localtime:
        # print "Server Time: ",
        # print localtime
        # # print "Local ctime: " + os.path.getctime(cwd + "/" + item)
        # print "Local Time: ",
        # print servertime
        # # print "Server ctime: " + os.path.getctime(serverpath + "/" + item)
        print ("[New Version Available!]"),
        cwd = os.getcwd()
        updater = os.path.abspath(cwd + r'\Configuration_Files\upit.exe')
        subprocess.Popen([updater])
        sys.exit(0)  
    else:
        print ("[No NEW Template Generator Version found]")          
except Exception as e:
    print ("Failed to get new version of the Template Generator!")
    print (e)

# updateScripts()

###This function makes sure the user has the most recent word xml template from the server
try:
    serverpath = getPath('serverconfig')
    # get path to current working directory
    cwd = os.getcwd()
    pcpath = os.path.abspath(cwd + '\Configuration_Files')
    bb=0
    i=0
    # Search through the word xml templates on the users computer
    for pcfn in os.listdir(pcpath):
        # Check to make sure you have the SAW word xml  and if it is get the time stamp on the file
        if pcfn.endswith(".xml"):
            pcfntime = os.path.getmtime(pcpath + "/" + pcfn)
            pcfname = pcfn
            # change bb to 1 to indicate that a SAW word xml  was found on users PC - this is used for console feedback
            bb=1
    # Search through the server configuration folder to see if there is a new word xml 
    print ("Searching for new Word Template..."),
    for svfn in os.listdir(serverpath):
        #check to make sure you have the SAW word xml  and then get the time stamp on the file
        if svfn.endswith(".xml"):
            svfntime = os.path.getmtime(serverpath + "/" + svfn)
            svfname = svfn
    # if a word xml  is present and the users file is up to date, do nothing and let the user know
    if bb == 1 and pcfntime >= svfntime:
        print ("[No NEW Word template found]")
    # if there is not a word xml  at all then copy a file
    elif bb == 0: 
        print ("[Word template missing!]")
        print ("Copying NEW Word XML template builidng block! DO NOT CLOSE THIS WINDOW!..."),
        #copy the new file
        copy(serverpath + "/" + svfname, pcpath)
        print ("[Complete]")
    # if there is a word xml  but it is older then delete the told one and then copy the new one 
    else:
        print ("[Found NEW Word template!]")
        print ('Copying NEW Word template! DO NOT CLOSE THIS WINDOW!...'),
        #remove the file on the local computer
        os.remove(pcpath + "/" + pcfname)
        #copy the new file
        copy(serverpath + "/" + svfname, pcpath)
        print ("[Complete]")
except Exception as e:
    print (e)
    print ("[Failed to copy Word XML templatee from the server!]")

### This checks the building block file on the server and replaces it if it has been modified ###
try:
    if getSettings('bbfolder') == "":
        pcpath = "C:/Users/" + getpass.getuser() +"/AppData/Roaming/Microsoft/Document Building Blocks/1033/15"
    else:
        pcpath = getSettings('bbfolder')
    serverpath = getPath('serverconfig')
    bb=0
    i=0
    # Search through the building blocks on the users computer
    for pcfn in os.listdir(pcpath):
        # Check to make sure you have the SAW building block and if it is get the time stamp on the file
        if pcfn.endswith(".dotx") and pcfn[:3]=='SAW':
            pcfntime = os.path.getmtime(pcpath + "/" + pcfn)
            pcfname = pcfn
            # change bb to 1 to indicate that a SAW building block was found on users PC - this is used for console feedback
            bb=1
        # else:
        #     # the file scanned is not a SAW building block
        #     pass
    # Search through the server configuration folder to see if there is a new building block
    print ("Searching for new common language building block..."),
    for svfn in os.listdir(serverpath):
        #check to make sure you have the SAW building block and then get the time stamp on the file
        if svfn.endswith(".dotx") and svfn[:3]=='SAW':
            svfntime = os.path.getmtime(serverpath + "/" + svfn)
            svfname = svfn
    # if a building block is present and the users file is up to date, do nothing and let the user know
    if bb == 1 and pcfntime >= svfntime:
        print ("[No NEW common language found]")
    # if there is not a building block at all then copy a file
    elif bb == 0: 
        print ("[Common language template missing!]")
        print ("Copying NEW common language builidng block! DO NOT CLOSE THIS WINDOW!..."),
        #copy the new file
        copy(serverpath + "/" + svfname, pcpath)
        print ("[Complete]")
    # if there is a building block but it is older then delete the told one and then copy the new one 
    elif pcfntime < svfntime:
        print ("[Found NEW common language!]")
        print ('Copying NEW common language building block! DO NOT CLOSE THIS WINDOW!...'),
        #remove the file on the local computer
        os.remove(pcpath + "/" + pcfname)
        #copy the new file
        copy(serverpath + "/" + svfname, pcpath)
        print ("[Complete]")
except Exception as e:
    print (e)
    print ("[Failed to copy common language from the server!]")

##### This checks the services json file and replaces it if it has been modified ###
try:
    # set the server path based on the configuration file but provide a default if they have deleted the configuration file
    try:
        serverpath = getPath('serverconfig')
    except:
        serverpath = '//saw-netapp2.saw.ds.usace.army.mil/Shared/Common/RG/CESAW-RG-A/David S/Developer/Configuration_Files/'
    ## get path to current working directory
    cwd = os.getcwd()
    pcpath = os.path.abspath(cwd + '\Configuration_Files')
    bb=0
    i=0
    # Search through the json services filetemplates on the users computer
    for pcfn in os.listdir(pcpath):
        # Check to make sure you have the SAW json services file and if it is get the time stamp on the file
        if pcfn == 'services.json':
            pcfntime = os.path.getmtime(pcpath + "/" + pcfn)
            pcfname = pcfn
            # change bb to 1 to indicate that a SAW json services file was found on users PC - this is used for console feedback
            bb=1
    # Search through the server configuration folder to see if there is a new json services file
    print ("Searching for new configuraton file..."),
    for svfn in os.listdir(serverpath):
        #check to make sure you have the SAW json services file and then get the time stamp on the file
        if svfn == 'services.json':
            svfntime = os.path.getmtime(serverpath + "/" + svfn)
            svfname = svfn
    # if a json services file is present and the users file is up to date, do nothing and let the user know
    if bb == 1 and pcfntime >= svfntime:
        print ("[No NEW configuration file found]")
    # if there is not a json services file at all then copy a file
    elif bb == 0: 
        print ("[Configuration file missing!]")
        print ("Copying NEW configuration file builidng block! DO NOT CLOSE THIS WINDOW!..."),
        #copy the new file
        copy(serverpath + "/" + svfname, pcpath)
        print ("[Complete]")
    # if there is a json services file but it is older then delete the told one and then copy the new one 
    else:
        print ("[Found NEW configuration file!]")
        print ('Copying NEW configuration file! DO NOT CLOSE THIS WINDOW!...'),
        #remove the file on the local computer
        os.remove(pcpath + "/" + pcfname)
        #copy the new file
        copy(serverpath + "/" + svfname, pcpath)
        print ("[Complete]")
except Exception as e:
    print (e)
    print ("[Failed to copy configuration file from the server!]")
