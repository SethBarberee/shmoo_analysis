#!/usr/bin/python3
#################################################################
#                   shmoo_dataset_to_excel.py                   #
#################################################################
#                                                               #
#   This script takes in a folder containing shmoo dataset and  #
#   converts its analysis into csv files and an excel           #
#                       spreadsheet.                            #
#                                                               #
#################################################################
# Version 0.2                                                   #
# By Shane Benetz                                               #
# Date: 09.15.2021                                              #
#################################################################
#################################################################
# Version 0.0 is first release 08.06.2021                       #
# Version 0.1 is last version with  08.25.2021                  #
# Version 0.2 is an update on edge cases for outputDir 09.15.21 #
#################################################################

import argparse
from datetime import datetime
import os
import glob
import sys
import re
from get_timing import get_timing, get_version
from shmoo_csv_to_excel import shmoo_csv_to_excel
from test_to_csv import test_to_csv
from process_shmoo_text import info_from_pathname

def shmoo_dataset_to_excel(inputDir, pCorner, outputDir, eqnset, list,\
     percentage, debug, comment, excluded) :  
    ''' Reads all shmoo plot files in a dataset directory
    and puts the Vmin/Vmax data and notes into a .xlsx file and .csv file
    Input:
        inputDir          --> folder to be read, default current
        process corners   --> test process corners to look for
        eqnset            --> where to find timing file
        outputDir         --> name of folder to write results to
        debug mode 0-3    --> turn on debugging print statements
        list              --> list all test folders being processed
        percentage        --> percentage from Vnom to look at
        comment           --> add path as comment in excel to bad tests
    Output:
        Complete dataset info in an excel format
    '''
    # check if input exists
    if not os.path.isdir(inputDir) :
        print('%s is not a directory' % inputDir)
        return
    # find correct output directory location
    try:
        if not (os.path.isdir(outputDir)) :
            os.makedirs(outputDir)
            print('Output folder created: ',os.path.abspath(outputDir))
        outputDir = os.path.abspath(outputDir)
        if not os.access(outputDir, os.W_OK) or not os.access(outputDir, os.R_OK):
            return print('Output directory not accessible')
    except: return print('Cannot use given output directory')


    # check if eqnset exists
    if eqnset != None :
        if not os.path.isfile(eqnset) :
            return print(eqnset + ' is not a file')
        else : # validate if its a real timing file
            EQname = os.path.basename(os.path.normpath(eqnset))
            with open(eqnset) as eq :
                lines = eq.readlines()
                if not ('EQNSET' in lines[0]) :
                    return print(EQname + ' is not a valid timing file')
    else :
        if os.path.isfile('./timingSpecs.txt') :
            eqnset = './timingSpecs.txt'
        elif os.path.isfile(os.path.join(outputDir,'timingSpecs.txt')):
            eqnset = os.path.join(outputDir,'timingSpecs.txt')
        else :
            get_timing(inputDir,outputDir)
            if not os.path.isfile(os.path.join(outputDir,'./timingSpecs.txt')) :
                return print('Please include timing file')
            eqnset = os.path.join(outputDir,'./timingSpecs.txt')
            print('Done!')
        try:
            with open(eqnset) as eq :
                lines = eq.readlines()
                if not ('EQNSET' in lines[0]) :
                    return print('Please provide a valid timingSpecs.txt file')
        except: return print('Please provide a valid timingSpecs.txt file')
    #get directories with process corner
    pCorner = pCorner.replace(' ', '')
    if len(pCorner) != 2 and pCorner != 'ALL':
        print('No such thing as %s, processing ALL instead' % pCorner)
        pCorner = 'ALL'
    paths = glob.glob('%s/*/'% inputDir) # list of all subdirectories
    testDirs = []
    pCorners = ['FF','SS','TT','FS','SF'] 
    for path in paths: #get all folders with process corner and temp in name
        pathName = os.path.abspath(path)
        info = info_from_pathname(pathName) 
        if (os.path.isdir(path)) and info[0] != None and info[1] != None :
            print("Checking path: " + path)
            if (pCorner == 'ALL') and any(pc in pathName for pc in pCorners) :
                if not any(dir in path for dir in excluded) : testDirs.append(path)
            elif (pCorner in pathName) :
                if not any(dir in path for dir in excluded) : testDirs.append(path)
            else : 
                paths.extend(glob.glob('%s/*/'% path)) # if its not, check subdirs
        else :
            paths.extend(glob.glob('%s/*/'% path)) # if its not, check subdirs
#    sys.exit()
    if len(testDirs) == 0 :
        return print('Please enter dataset folder containing %s tests.' % pCorner)

    #check percentage 
    try:
        decimal = percentage
        if '%' in percentage :
            decimal = float(re.sub('[^0-9-]','',percentage))/100
        elif -1.00 <= float(percentage) <= 1.00:
            decimal = float(percentage)
        elif 1 < float(percentage) < 100 or -100 < float(percentage) < -1 :
            decimal = float(percentage)/100
        else: return print('Cannot read given percentage: ' + percentage)
    except:
        return print('Cannot read given percentage: ' + percentage)
    percentageOut = ('%f' % (decimal*100)).rstrip('0').rstrip('.') + '%'

    # getting the product name from parent directories
    fullPath = os.path.abspath(os.path.realpath(inputDir))
    parents = fullPath.split(os.sep)
    productDir = os.sep
    ctr = 0
    for parent in parents : # add parent dir until product dir is found
        if ctr == 3 : break
        productDir = os.path.join(productDir,parent)
        if ctr >0 : ctr +=1
        if parent == 'designs' : ctr +=1

    # make the correct output file
    pName = None
    grandparent = os.path.basename(os.path.abspath(os.path.join(productDir,'..','..')))
    if  grandparent == 'designs' :
        pName = os.path.basename(os.path.normpath(productDir))
    if pName == None : 
        pName = 'dataset'
    if pCorner == 'ALL' :
        outputFile = '_'.join([pName,percentageOut,'notes.xlsx'])
    else : 
        outputFile = '_'.join([pName,pCorner,percentageOut,'notes.xlsx'])
    outputFile = os.path.join(outputDir, outputFile)
    i = 0
    while os.path.isfile(outputFile): # add nuber version if repeat
        i+=1
        outputFile = outputFile[:outputFile.find('_notes')] + '_notes_'+str(i)+'.xlsx'

    # get all data into CSV files
    if len(testDirs) == 0 : return print('No Test Directories Found')
    print('\nWorking on CSV conversion ', end='', flush=True) 
    printed = False
    removed = []
    outputs = ['']
    csvFiles = []
    time = datetime.now().strftime('%m-%d-%Y_%Hh%Mm%Ss')
    for folder in testDirs :
        if list: 
            if printed == False: print()
            print(folder); printed = True
        try:
            folderName = os.path.normpath(folder)
            process = None
            for pc in pCorners :
                if pc in folderName :
                    process = pc
            if process == None : break
            if '-' in percentage :
                outputCSV = os.path.join(outputDir, \
                    'Vmin_'+ percentageOut.replace('%','p')+ '_'+ process + '.csv')
            else: 
                outputCSV = os.path.join(outputDir, \
                    'Vmax_+'+ percentageOut.replace('%','p')+ '_'+ process + '.csv')
            # add Vmin/Vmax CSVs to list
            if not outputCSV in csvFiles: csvFiles.append(outputCSV)
            if not os.path.isfile(outputCSV) :
                removed.append(outputCSV)
            if (not outputCSV in removed) and os.path.isfile(outputCSV) :
                os.remove(outputCSV)
                removed.append(outputCSV)
            # add Tmin/Tmax CSVs to list
            tOutputCSV = outputCSV.replace('Vm','Tm')
            if not tOutputCSV in csvFiles: csvFiles.append(tOutputCSV)
            if not os.path.isfile(tOutputCSV) :
                removed.append(outputCSV)
            if (not tOutputCSV in removed) and os.path.isfile(tOutputCSV) :
                os.remove(tOutputCSV)
                removed.append(tOutputCSV)
            out = test_to_csv(folder,outputCSV,eqnset,None,None,percentage,debug,time)
            outString = str(out).splitlines()
            if len(outString) > 0 :
                for line in outString :
                    if not str(line) in outputs:
                        if len(outputs) == 1: print('')
                        outputs.append(line)
                        print(line, flush=True)
                        printed = True
        except KeyboardInterrupt : 
            sys.exit(' Keyboard Interruption: Processes killed')
        except Exception as e:
            print(e)
            print('\nCannot Process ' + folder)  
            printed = True
    if not printed: print('- Done!')
    else : print('Done!')

    # convert recently created CSVs
    finalCSVs = []
    for csv in csvFiles:
        if pCorner == 'ALL' :
            finalCSVs.append(csv)
        elif pCorner in csv :
            finalCSVs.append(csv)
    newInfo = True
    last = False
    for i in range(len(finalCSVs)) :
        if i == len(finalCSVs)-1 : last = True
        shmoo_csv_to_excel(finalCSVs[i],outputFile,outputDir,False,newInfo,comment,last)
        newInfo = False
    if os.path.isfile(outputFile):
        print('Analysis notes saved in: ' + '\x1b[0;30;43m' +\
             os.path.abspath(outputFile) + '\x1b[0m')

if __name__=='__main__' :
    for arg in range(0,len(sys.argv)) :
        if '%' in sys.argv[arg] : # in case the percentage sign messes up argparse
            sys.argv[arg] = sys.argv[arg].replace('%','') 
    parser = argparse.ArgumentParser(description='Transform Shmoo Dataset into Excel',\
        formatter_class=argparse.RawTextHelpFormatter,\
        epilog = 'usage examples:\n'\
        '  shmoo_dataset_to_excel -i char -o analysis/ -c FF -l -n -p -6%\n\n'\
        '  shmoo_dataset_to_excel -i char -o analysis/ -c FF -l -n -p -6% -x snFF01_25C (to exclude a serial number at a temp)\n\n'\
        '  shmoo_dataset_to_excel -i char -eq ../timingSpecs.txt -d 3 -x junk \n')
    parser.add_argument('-v', '-V', '--version', dest='version', action='store_true',\
        default=False, help='get version of script and exit')
    parser.add_argument('-i', '--input', dest='inputDir', default='.', \
        help='input folder location. DEFAULT current folder')
    parser.add_argument('-c', '--corner', dest='pCorner', default='ALL', \
        choices=['ALL', 'FF', 'TT', 'SS', 'FS', 'SF'],\
        help='name of process corner to assess. DEFAULT = ALL')
    parser.add_argument('-o', '--output', dest='outputDir', default='.', \
        help='output folder path. DEFAULT current folder.')
    parser.add_argument('-eq', '--timing', dest='eqnset', default=None, \
        help='spec periods file location; default checks if in current folder' +\
            '\n(usually called timingSpecs.txt. use get_timing command to get this file)')
    parser.add_argument('-L', '--list', dest='list', action='store_true', \
        default=False, help='list test folders being processed')
    parser.add_argument('-p', '--percent', dest='percentage', default='-5%%', \
        help='percentage from Vnom. DEFAULT = -5%%')
    parser.add_argument('-x', '--exclude', nargs='+',dest='exclude', default=[], \
        help='names of files/directories to be excluded')
    parser.add_argument('-n', '--notes', dest='comment', action='store_true', \
        default=False, help='add path in an excel note for all problematic tests') 
    parser.add_argument('-d', '--debug', dest='debug', type=int, default=0, \
        choices=range(0,4), metavar='[0-3]', help='debug mode 0-3.\n' \
            '0 = No Messages\n' '1 = Error Mesages\n' '2 = Plots w/ fails or holes\n' \
                '3 = All Bad Plots w/ Notes')
    args = parser.parse_args()
    if args.version :
        version = get_version(sys.argv[0])
        print('Version: ' + version)
        sys.exit()
    try: 
        shmoo_dataset_to_excel(args.inputDir, args.pCorner, args.outputDir, args.eqnset, \
        args.list,args.percentage,args.debug,args.comment,args.exclude)
    except KeyboardInterrupt :
        print('\nKeyboard Interrupt: Process Killed')
        sys.exit()
   
