#!/usr/bin/python3
#################################################################
#                         get_voltage                    		#
#################################################################
#                                                               #
#   This script takes in a folder/file and goes through all the #
#   subfolders or through file to get the spec vnoms for each   #
#   Outputs file: voltageSpecs.txt                              #
#                                                               #
#################################################################
# Version 0.1                                                   #
# By Shane Benetz                                               #
# Date: 12.29.2021                                              #
#################################################################
#################################################################
# Version 0.0 is first release 12.29.2021                       #

#################################################################

version = 0.0

import argparse
import sys
import os
import subprocess
import re

def get_voltage_specs(input,outputDir) : 
    # finds all the levels folders under given folder
    fullPath = os.path.abspath(os.path.join(os.path.realpath(input),os.pardir))
    if not ('us' in fullPath or 'ops' in fullPath) : return # only our linux server
    parents = fullPath.split(os.sep)
    tempDir = os.sep
    ctr = 0
    for parent in parents :
        if ctr == 3 : break
        tempDir = os.path.join(tempDir,parent)
        if ctr > 0 : ctr +=1
        if parent == 'designs' : ctr +=1
    if ctr == 0 : return
    # find all specsets with nominal name and dont repeat and get only voltages
    cmd = 'find %s -type d -name \"levels\" |& grep -v \"Permission denied\" | ' % tempDir \
        + 'xargs grep -ihr -A 45 \"SPECSET \" | grep -A 45 \"nom\" | sort -u | grep -o '\
        + '"^V.*\"' 
    try : 
        levels= subprocess.check_output(cmd, shell = True).decode('UTF-8')
    except : 
        print('Issue getting level specs',flush=True)
        return
    if len(levels) == 0 :return
    if not os.path.isdir(outputDir) : os.makedirs(outputDir)
    vFile = os.path.join(outputDir,'voltageSpecs.txt')
    with open(vFile,'w') as specV :
        specV.write(levels) 
        
if __name__ == '__main__' :
    parser = argparse.ArgumentParser(description='Get voltages for a project', \
        formatter_class = argparse.RawTextHelpFormatter, epilog = 'usage examples:\n'\
        '  get_voltage -i designs/fadu/soyang -o ./outputFolder/ \n\n')
    parser.add_argument('-v', '-V', '--version', dest='version', action='store_true',\
        default=False, help='get version of script and exit')
    parser.add_argument('-i', '--input', dest='input', default=None, \
        help='path to folder or file that you want voltage for. only works on linux '\
            'file structure. will return all voltages for project in one file')
    parser.add_argument('-o', '--output', dest='outputDir', default='.', \
        help='output folder path. creates output path if DNE. DEFAULT current folder')
    args = parser.parse_args()
    if args.version :
        print('Version: ' + str(version))
        sys.exit()
    get_voltage_specs(args.input, args.outputDir)
