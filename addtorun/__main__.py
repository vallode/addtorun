#!/usr/bin/env python

import _winreg, argparse, os, colorama

from _winreg import *
from colorama import init, Fore, Back, Style
from os import listdir, stat

def main():
    
    init(autoreset=True); # Colorama, reset it's style/color options at the end of everyline. Just some *\*/**AESTHETIC**\*/*
    path = listdir('.') # Locate the directory in which this is being run from

    # Can not think how to do this better so placeholder numbers
    fileStatPath = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    fileStatName = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    fileStatPathFolder = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    fileNum = 0
    choiceNum = 1

    #*\*/**AESTHETIC**\*/*
    print Fore.YELLOW + Style.BRIGHT + '<AddToRun>'
    print Style.BRIGHT + 'Directory:'
    
    #Run through all files in path and list the ones that end with .exe
    for files in path:
        if path[fileNum].endswith('.exe'):
            print 'Option: ' + str(choiceNum) + ' ||', path[fileNum]
            fileStatPath[choiceNum] = os.path.abspath(path[fileNum])
            fileStatName[choiceNum] = os.path.basename(path[fileNum])
            fileStatPathFolder[choiceNum] = fileStatPath[choiceNum].replace(fileStatName[choiceNum], '') 
            print 'Path: ' + fileStatPath[choiceNum]
            print 'Name: ' + fileStatName[choiceNum] + '\n'
            fileNum = fileNum + 1 
            choiceNum = choiceNum + 1
        else:
            fileNum = fileNum + 1
            
    #Ask the user to input a valid integer from the list of options
    while True:
        try:
            choiceNum = int(raw_input('Which program would you like to add to the Run dialog?\n'))
            break
        except (ValueError):
            print Fore.RED + Style.BRIGHT + 'Not a valid option!'
    
    while True:
        try:
            fileStatName[choiceNum] = str(raw_input('What would you like to call the shortcut?'))
            break
        except (ValueError):
            print Fore.RED + Style.BRIGHT + 'Not a valid option!'
    while True:
        if ('.' in fileStatName[choiceNum]):
            fileStatName[choiceNum] = str(raw_input('You shortcut should not contain a full stop "." to avoid errors\n') + '.exe')
            break


    #This is opening your local machine registry at the App Paths
    #where the run dialog shortcuts are located
    
    print fileStatPath[choiceNum]
    print fileStatName[choiceNum]

    Key = OpenKey(_winreg.HKEY_LOCAL_MACHINE,
                     'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\\',0,
                     _winreg.KEY_ALL_ACCESS)

    #Sets the FULL PATH of the shortcut
    _winreg.SetValue(Key,
                     fileStatName[choiceNum],
                     _winreg.REG_SZ,
                     fileStatPath[choiceNum])

    Key = OpenKey(_winreg.HKEY_LOCAL_MACHINE,
                     'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\\' + fileStatName[choiceNum] + '\\',0,
                     _winreg.KEY_ALL_ACCESS)

    #Sets the FOLDER of the path in the registry
    _winreg.SetValueEx(Key,
                       'Path',
                       0,
                       _winreg.REG_SZ,
                       fileStatPathFolder[choiceNum])

    _winreg.CloseKey(Key)    
        
if __name__ == '__main__':
    main()
