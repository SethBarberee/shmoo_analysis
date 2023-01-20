#!/usr/bin/python3
#################################################################
#                         get_timing                    		#
#################################################################
#                                                               #
#   This script takes in a folder/file and goes through all the #
#   subfolders or through fileto get the spec period for each   #
#   equation set (eg. EQNSET 1). Outputs file: timingSpecs.txt  #
#                                                               #
#################################################################
# Version 0.2                                                   #
# By Shane Benetz                                               #
# Date: 09.15.2021                                              #
#################################################################
#################################################################
# Version 0.0 is first release 08.06.2021                       #
# Version 0.1 updated path reading and usage info 08.19.2021    #
# Version 0.2 is an update on edge cases for outputDir 09.15.21 # 
#################################################################

import argparse
import sys
import os
import subprocess
import re

def get_timing(input,outputDir):
    #get input 
    if input == None :
        fullPath = os.path.abspath('.')
        parents = fullPath.split(os.sep)
        timingDir = os.sep
        ctr = 0
        for parent in parents :
            if ctr == 3 : break
            timingDir = os.path.join(timingDir,parent)
            if ctr >0 : ctr +=1
            if parent == 'designs' : ctr +=1 #proj name 2 dirs past designs dir 
        input = timingDir
    elif os.path.basename(input).endswith('.tim') and not os.path.isfile(input) :
       print(input + ' is not a valid folder'); sys.exit()
    else : timingDir = os.path.realpath(input)
     
    if not os.path.basename(input).endswith('.tim') and not os.path.isdir(timingDir):
        print(timingDir + ' is not a valid folder')
    # find correct output directory location
    try:
        if not (os.path.isdir(outputDir)) :
            os.makedirs(outputDir)
            print('Output folder created: ',os.path.abspath(outputDir))
        outputDir = os.path.abspath(outputDir)
        if not os.access(outputDir, os.W_OK) or not os.access(outputDir, os.R_OK):
            return print('Output directory not accessible')
    except: return print('Cannot use given output directory')
    print('\rSearching for timing files...',end='\r')
    if os.path.basename(input).endswith('.tim') : #if timing file only provided
        tDirsFinal = [input]
    else:
        # finds all the timing folders under given folder
        cmd = 'find '+timingDir+' -type d -name \"timing\" |& grep -v \"Permission denied\"'
        try:
            timingDirsList = subprocess.check_output(cmd, shell = True).decode('UTF-8')
        except subprocess.CalledProcessError as e : 
            pass
            fullPath = os.path.abspath(input)
            parents = fullPath.split(os.sep)
            timingDir = os.sep
            ctr = 0
            for parent in parents :
                if ctr == 3 : break
                timingDir = os.path.join(timingDir,parent)
                if ctr >0 : ctr +=1
                if parent == 'designs' : ctr +=1
            cmd = 'find '+timingDir+' -type d -name \"timing\" |& grep -v \"Permission denied\"'
            try : 
                timingDirsList = subprocess.check_output(cmd, shell = True).decode('UTF-8')
            except : 
                print('Issues finding valid timing folder',flush=True)
                sys.exit()  
        timingDirsList = timingDirsList.split('\n')
        tDirsFinal = []
        for tF in timingDirsList:
            if os.path.isdir(tF) : tDirsFinal.append(tF) 
    print('Looking for specs in timing files:                ')
    for dir in tDirsFinal :
        print(dir,flush=True)
    output = ''
    tOutput = ''
    if len(tDirsFinal) == 0 :
        if input == None : print('Please provide timing folder')
        else : print('Provided directory does not contain timing folders')
        sys.exit()
    else:  
        for tDir in tDirsFinal :
            # gets chunk of 8 lines that contain eqnset # and period 
            try:
                output = output + subprocess.check_output('grep -ihr -A 8 \'eqnset\' ' +\
                os.path.abspath(tDir), shell = True).decode('UTF-8')
            except:pass
            try: 
                tOutput = tOutput + subprocess.check_output('grep -ihr -A 2 \"TIMINGSET\" ' +\
                os.path.abspath(tDir), shell = True).decode('UTF-8')
            except:pass
    lines1 = output.splitlines()
    lines2 = tOutput.splitlines()
    try :
        with open(os.path.join(outputDir,'timingSpecs.txt'), 'w') as out :
            ctr = 0
            lines_seen = {}
            for line in lines1 :
                line = str(line)
                if line.find('EQNSET') >=0 and ctr < len(lines1) - 8:
                    period = 0
                    eqnset = int(line[line.find('EQNSET')+ 6 : line.find('\"')-1])
                    six = str(lines1[ctr +6])
                    seven = str(lines1[ctr + 7])
                    eight = str(lines1[ctr + 8])
                    if six.lower().find('period') >= 0 and \
                        (six.find('ns') > 0 or six.find('\#') > 0) : # if 6th line has period
                        period = str(re.sub('[^0-9.]', '', six[six.find(' '):]))
                    elif seven.lower().find('period') >= 0 and \
                        (seven.find('ns') > 0 or seven.find('#') > 0) : #if 7th line has period
                        period = str(re.sub('[^0-9.]', '', seven[seven.find(' '):]))
                    elif eight.lower().find('period') >= 0 and \
                        (eight.find('ns') > 0 or eight.find('#') > 0) : #if 8th line has period
                        period = str(re.sub('[^0-9.]', '', eight[eight.find(' '):]))
                    try:
                        if len(period) == 0 or float(period) == 0 :
                            raise Exception
                    except:
                        ctr+=1
                        continue 
                    # add period to correct eqnset if its valid (stored in dictionary)
                    if lines_seen.get(eqnset) and len(period) > 0 :
                        if lines_seen.get(eqnset).find(str(period+',')) < 0 :
                            temp = lines_seen[eqnset]
                            lines_seen[eqnset] = temp + str(period) +', '
                    else:
                        lines_seen[eqnset] = period + ', '
                ctr +=1
            # sort eqnset by number and write to output file
            for key in lines_seen:
                out.write('EQNSET' + str(key) + ', ' + str(lines_seen[key])+ '\n')
            
            ctr = 0
            lines2_seen = {}
            for line in lines2 :
                line = str(line)
                if line.find('TIMINGSET') >=0 and ctr < len(lines1) - 2:
                    period = 0
                    timset = line[line.find('TIM'):line.rfind('\"')+1].replace('FRC_','')
                    nextLine = lines2[ctr+1]
                    if nextLine.lower().find('period') >= 0 and \
                        (nextLine.find('ns') > 0 or nextLine.find('#') > 0) : 
                        period = str(re.sub('[^0-9.]', '', nextLine[nextLine.find(' '):]))
                    try:
                        if len(period) == 0 or float(period) == 0 :
                            raise Exception
                    except:
                        ctr+=1
                        continue 
                    # add period to correct eqnset if its valid (stored in dictionary)
                    if lines2_seen.get(timset) and len(period) > 0 :
                        if lines2_seen.get(timset).find(str(period+',')) < 0 :
                            temp = lines2_seen[timset]
                            lines2_seen[timset] = temp + str(period) +', '
                    else:
                        lines2_seen[timset] = period + ', '
                ctr +=1
            # sort eqnset by number and write to output file
            for key in lines2_seen:
                out.write(str(key) + ', ' + str(lines2_seen[key])+ '\n')
            print('Timing file location: '+ os.path.join(outputDir,'timingSpecs.txt'))
    except PermissionError:
        print('Please provide output folder with write permission')

def get_version(filename) :
    '''Finds first instance of version listed in file header'''
    version = 'Cannot Find Version'
    if not os.path.exists(filename) : return version
    try:
        with open(filename,'r') as file :
            contents = file.read()
            # top most version line
            versionIndex = contents.lower().find('version')
            versionLine = contents[versionIndex:contents[versionIndex:].find('#')+versionIndex]
            version = versionLine[versionLine.lower().find('version') + 7:].replace(' ','')
    except: pass
    return version
if __name__ == '__main__' :
    parser = argparse.ArgumentParser(description='Analyze a shmoo plot files', \
        formatter_class = argparse.RawTextHelpFormatter, epilog = 'usage examples:\n'\
        '  get_timing -i designs/fadu/soyang -o ./timingFolder/ \n\n'\
        '  get_timing -i soyang/pgm/main/Soyang_FT/timing/soyang_A06.tim\n')
    parser.add_argument('-v', '-V', '--version', dest='version', action='store_true',\
        default=False, help='get version of script and exit')
    parser.add_argument('-i', '--input', dest='input', default=None, \
        help='path to folder that is or contains timing folder(s) (.../timing/...)'\
            ' or the path to a testflow timing file (*.tim)')
    parser.add_argument('-o', '--output', dest='outputDir', default='.', \
        help='output folder path. creates output path if DNE. DEFAULT current folder')
    args = parser.parse_args()
    if args.version :
        version = get_version(sys.argv[0])
        print('Version: ' + version)
        sys.exit()
    get_timing(args.input, args.outputDir)
