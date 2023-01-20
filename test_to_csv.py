#!/usr/bin/python3
#################################################################
#                         test_to_csv                    		#
#################################################################
#                                                               #
#   This script takes in a folder containing shmoo data from    #
#   one testflow (eg. TT02_25C) and process the most recent     #
#   files in it for each sub-test and turns the data into CSV   #
#                                                               #
#################################################################
# Version 0.2                                                   #
# By Shane Benetz                                               #
# Date: 09.15.2021                                              #
#################################################################
#################################################################
# Version 0.0 is first release 08.06.2021                       #
# Version 0.1 is last version before adding Tmin 08.25.2021     #
# Version 0.2 is an update on edge cases for outputDir 09.15.21 #
#################################################################
import glob
import sys
import re
import os.path
import argparse
from datetime import datetime
from get_timing import get_timing, get_version
from process_shmoo_text import process_shmoo_text, info_from_pathname

def test_to_csv(input, outputFile, eqnset, device, temp, percentage, debug, time) :
    ''' Reads all shmoo plot files in a complete device + temp test directory
    and puts the Vmin/Vmax data and notes into a .csv file
    Inputs:
        input       --> folder to be read, default current
        outputFile  --> .csv name of file to write results to
        eqnset      --> file to find spec period in
        device      --> device name
        temp        --> testing temperature
        debug       --> turn on debugging print statements (mode 0-3)
    Output:
        Data for one shmoo folder, either device + temp folder or just one
        pattern, to .csv file
    '''
    # check if input exists
    if not os.path.isdir(input) :
        return print('%s does not exist' % input)
    paths = glob.glob('%s/*/'% input) # list of all subdirectories

    # check output directory
    if outputFile != None :
        outputDir = os.path.abspath(os.path.join(outputFile, os.pardir))
    else: outputDir = '.'
    if outputFile != None and not (outputFile.endswith('.csv')) :
        return print('\nPlease provide a .csv output file name')
    elif outputFile != None and not os.path.isdir(outputDir) : 
        os.makedirs(outputDir)
    if not os.access(outputDir, os.W_OK) or not os.access(outputDir,os.R_OK) : 
        print('Cannot write to given output Directory'); sys.exit()
    #check eqnset
    if eqnset != '.' : #check the given timing spec file
        if not os.path.isfile(eqnset) :
            return print('\n' + eqnset + ' is not a file')
        if eqnset.endswith('.tim'): 
            if outputFile != None : get_timing(eqnset,outputDir)
            eqnset = os.path.join(outputDir,'timingSpecs.txt')
        with open(eqnset) as eq :
            lines = eq.readlines()
            if not ('EQNSET' in lines[0]) :
                return print('\n' + eqnset + ' is not a valid timing file')

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

    # get temp and device if not given
    if device == None or temp == None :
        fullPath = os.path.abspath(input)
        getInfo = info_from_pathname(fullPath)
        if temp == None : temp = getInfo[0]
        if device == None :device = getInfo[1]
    if temp == None or device == None:
        if __name__ != '__main__' :
            if debug > 0 :
                return ['Cannot find Temp/Device for test folder' + input, None]
        else:
            if temp == None : 
                print('Please provide temperature (-T temp)')
            if device == None : 
                print('Please provide device name (-D name)')
            return
    
    # find correct output location
    if outputFile == None :
        if '-' in percentage : 
            outputFile = 'Vmin('+ percentageOut+ ')_' + device + '_' + temp + '.csv'
        else : 
            outputFile = 'Vmax(+'+ percentageOut+ ')_' + device + '_' + temp + '.csv'
    orig_stdout = sys.stdout
    if debug > 1 :
        f = open(os.path.join(outputDir,'debug' + time + '.txt'), 'a')
        sys.stdout = f
    files = []
    output = ''
    count = 0
    for pathName in paths :
        # if there are subdirectories in folder
        if os.path.isdir(pathName) and len(glob.glob('%s/*/' % pathName)) > 0 :
            paths.extend(glob.glob('%s/*/' % pathName))
            continue
        # get all .txt from the folder
        files = glob.glob(os.path.join(pathName,'*.txt'))
        for file in files :
            count+=1
            try :
                out = process_shmoo_text([file],outputFile,eqnset,None,device,temp,\
                    percentage,debug)
                if str(out) not in output and str(out) != 'None': 
                    output = output + str(out)
            except Exception as e :
                 print('Error reading ' + file + ' ' + str(e))
    try: # sort lines
        with open(outputFile, 'r') as oldFile :   
            content = oldFile.readlines()
            lines = [content[0]] + sorted(content[1:])
        with open(outputFile,'w') as newFile :
            newFile.write(''.join(lines))
        tOutputFile = os.path.join(os.path.dirname(outputFile),\
            os.path.basename(outputFile).replace('Vm','Tm'))
        with open(tOutputFile, 'r') as oldFile :   
            content = oldFile.readlines()
            lines = [content[0]] + sorted(content[1:])
        with open(tOutputFile,'w') as newFile :
            newFile.write(''.join(lines))
    except Exception as e: print(e);pass
    if debug > 1 :
        sys.stdout = orig_stdout
        f.close()
    if os.path.isfile(outputFile) and __name__ == '__main__': 
        print('CSV Location: ' + '\x1b[0;30;43m' +\
             os.path.abspath(outputFile) + '\x1b[0m',end='')
    return output

if __name__ == '__main__':
    for arg in range(0,len(sys.argv)) :
        if '%' in sys.argv[arg] : # in case the percentage sign messes up argparse
            sys.argv[arg] = sys.argv[arg].replace('%','') 
    parser = argparse.ArgumentParser(description='Get data from one testflow\'s' + \
        ' shmoo text folder and transform it into a .csv', formatter_class = \
            argparse.RawTextHelpFormatter,\
            epilog = 'usage examples:\n'\
        '  test_to_csv -i char/Char/snTT2_25C/ -eq char/timingSpecs.txt -d 1\n\n'\
        '  test_to_csv -i char/Char/snFF3_-5C/ -d 3 -p -4% -o CSV/FF.csv\n')
    parser.add_argument('-v', '-V', '--version', dest='version', action='store_true',\
        default=False, help='get version of script and exit')
    parser.add_argument('-i', '--input', dest='inputDir', default=None, \
        help='required input folder location')
    parser.add_argument('-o', '--output', dest='outputFile', default=None, \
        help='output file path(.csv only w/ Vmin/Vmax in name) DEFAULT device_temp.csv')
    parser.add_argument('-eq', '--timing', dest='eqnset', default='.', \
        help='spec periods file location; default checks if in current folder' +\
            '\nusually called timingSpecs.txt(use get_timing command to get this file)')
    parser.add_argument('-D', '--device', dest='device', default=None, \
        help='device name (eg. TT02)')
    parser.add_argument('-T', '--temperature', dest='temp', default=None, \
        help='testing temperature (eg. 25C)')
    parser.add_argument('-p', '--percent', dest='percentage', default='-5%%', \
        help='percentage from Vnom. DEFAULT = -5%%')
    parser.add_argument('-d', '--debug', dest='debug', type=int, default=0, \
        choices=range(0,4), metavar='[0-3]', help='debug mode [0-3]\n' \
            '0 = No Messages\n' '1 = Error Mesages\n' '2 = Some Bad Plots\n' \
                '3 = All Bad Plots w/ Notes')
    args = parser.parse_args()
    time = datetime.now().strftime('%m-%d-%Y_%Hh%Mm%Ss')
    if args.version :
        version = get_version(sys.argv[0])
        print('Version: ' + version)
        sys.exit()
    if args.inputDir == None :
        parser.print_usage()
        print('Input directory is required. Please include using -i')
        sys.exit()
    out = test_to_csv(args.inputDir,args.outputFile,args.eqnset,args.device,\
        args.temp,args.percentage,args.debug,time)
    if out != 'None' : print(out)




