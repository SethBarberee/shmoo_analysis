#!/usr/bin/python3
#################################################################
#                         process_shmoo_text                    #
#################################################################
#                                                               #
#    This script takes in a list of shmoo text files and        #
#   analyzes them to output their parameters, minimum passing   #
#   voltage, any issues that are found with the graphs          #
#                                                               #
#################################################################
# Version 0.3                                                   #
# By Shane Benetz                                               #
# Date: 10.25.2021                                              #
#################################################################
#################################################################
# Version 0.0 is first release 08.06.2021                       #
# Version 0.1 updates path handling + usage examples 08.20.2021 #
# Version 0.2 fixes error handling that was ingonring good files#
# Version 0.3 fixes errors from flipped axises on shmoo         #
#################################################################

import os
import argparse
import re
import sys
from datetime import datetime
from get_timing import get_timing, get_version
import subprocess

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

def params(filename,outputDir) : 
    ''' Helper module to get parameters from header of shmoo file 
    Inputs:
        filename - file to be read for parameters

    Outputs:
        (0)patternName - title of shmoo file
        (1)xMin, (2)xMax, (3)xStep - the min, max, and step size of the X-axis 
        (4)yMin, (5)yMax, (6)yStep - the min, max, and step size of the Y-axis 
        (7)Vnom - spec voltage
        (8)eqNum - timing equation set to find period 
    '''
    with open(filename,'r') as file :
        patternName = domain = range = xStep = yStep = None
        ctr = 0; Vnoms = []; specName = None; VnomFile = 0
        contents = file.readlines()
        for line in contents :
            if (patternName == None) and (line.find('Title') >= 0): 
                pLine = contents[ctr]
                #finds the title of the plot
                patternName = pLine[pLine.find(':') + 2 : pLine.find('\n')] 
            # X-axis descriptions starts
            elif (domain == None) and (line.find('X-Axis') >= 0) :
                domainLine = contents[ctr + 3] # X-range 3 lines after
                domain = domainLine[domainLine.find('[') : ]
                xStepLine = contents[ctr + 6] # X-step 6 lines after
                xStep = xStepLine[xStepLine.find(':') + 2 : xStepLine.find('\n')] 
                # get period EQNSET from spec line
                specLine = contents[ctr + 7] # timing spec 9 lines after
                eqNum = re.sub('[^0-9]', '', specLine[:specLine.rfind('.')]) # gets spec timing num
            #Y-axis descriptions starts
            elif (range == None) and (line.find('Y-Axis') >= 0) :
                yrangeLine = contents[ctr + 3] #Y-range 3 lines after                   
                range = yrangeLine[yrangeLine.find('[') : ]    
                yStepLine = contents[ctr + 6] #Y-step 6 lines after
                yStep = yStepLine[yStepLine.find(':') + 2 : yStepLine.find('\n')]
                #get spec voltage from description line or from title
                descriptionLine = contents[ctr + 9] #description 9 lines after
                voltageNum = re.sub('[^0-9]', '', descriptionLine) #get spec voltage
                try:
                    VnomFile = voltageNum[0] + '.' + voltageNum[1:] # add decimal point
                except:
                    try:
                        voltageNum = re.sub('[^0-9]', '', \
                            patternName[patternName.find('VDD') + 3 : ]) 
                        VnomFile = voltageNum[0] + '.' + voltageNum[1:] # add decimal point
                    except:
                        pass
                # get voltage from level name
                specName = descriptionLine[descriptionLine.find('.')+1:]
                specName = specName.replace(' ','').strip('\n')
            # end if all the values are filled out
            elif patternName != domain != range != xStep != yStep != None :
                break
            ctr += 1
        #get vnom from voltagespecs.txt file
        vFile = os.path.join(outputDir,'voltageSpecs.txt') 
        if not os.path.isfile(vFile):
            get_voltage_specs(filename,outputDir)
        if os.path.isfile(vFile) and specName:
            with open(vFile,'r') as specs:
                lines = specs.readlines()
                for line in lines :
                    if line.find(specName + ' ') > -1:
                        # lines look like this: 
                        # VDDQ_DDR         1.1                [  V] #1.2V or 1.1V
                        Vnomtemp = re.sub('[^0-9.]', '', line[line.find(' '):line.find('[')])
                        try:
                            float(Vnomtemp)
                            Vnoms.append(float(Vnomtemp))
                        except:
                            pass
        
        # get max and min from x&y ranges    
        xMin = domain[1 : domain.find('...') - 1].strip(' ')
        xMax = domain[(domain.find('...') + 4) : domain.find(']') ].strip(' ')
        flippedX = False
        if float(xMin) > float(xMax):
            temp = xMin
            xMin = xMax
            xMax = temp
            flippedX = True
        
        yMin = range[1 : domain.find('...') - 1].strip(' ')
        yMax = range[domain.find('...') + 3 : range.find(']')].strip(' ')
        if float(yMin) > float(yMax):
            temp = yMin
            yMin = yMax
            yMax = temp
        #if >1 voltage in voltageSpecs.txt, get closest to midpoint of graph
        midVoltage = float((float(yMax) - float(yMin))/2 + float(yMin)*1.02)
        if len(Vnoms) == 0 :
            Vnom = VnomFile    
        else:
            closest = min(Vnoms, key = lambda i: abs(i-midVoltage))
            Vnom = closest
        if not(float(yMin) < float(Vnom) < float(yMax)):
            try: # move deimal point over 1 if its not in range
                # try to move decimal
                Vnom = str(Vnom)
                Vnom = Vnom[0]+Vnom[2]+Vnom[1]+Vnom[3:]
                Vnom = float(Vnom)
                if not(float(yMin) < float(Vnom) < float(yMax)):
                    raise Exception
            except Exception as e:
                return '\nVnom not found for '+ filename
        try: #get sig figs
            sigFigs = len(yMax[yMax.find('.')+1:]) #sig figs from parameter vMax
            Vnom = ('%.'+str(sigFigs)+'f') % float(Vnom)
        except: pass
        return [patternName, xMin, xMax, xStep, yMin, yMax, yStep, Vnom, eqNum, flippedX]
      
def passing_range(line, periodStepSize, minPeriod, specPeriod, flippedX) :
    ''' Finds the range of passing periods for one voltage line in plot
    Inputs:
        line            - voltage pass/fail plot line
        periodStepSize  - distance between points by period
        minPeriod       - starting period of graph (x_min)
    Output:
        4 item list -
        (0) passing or not 
        (1) first passing period - start of valid period range
        (2) last passing period - end of valid period range
        (3) holes - if holes exist in the range
    '''
    passing = False
    if line.count('.') == 0 :
        return [passing,0,0,False]
    else :
        holes = False
        pf = line[line.find('|')+ 2: ].strip('\n')
        if flippedX:
            periodStepSize = abs(periodStepSize)
            pf = pf[::-1]
        periodIndex = int(round((specPeriod - minPeriod),4)/periodStepSize)
        if pf[periodIndex] == '.' : passing = True
        if passing : firstIndex = pf[:periodIndex+1].rfind('*') + 1 #last fail +1 before spec period
        else : firstIndex = pf.find('...') # find first three passes in a row
        firstPass = minPeriod + firstIndex * periodStepSize
        if firstIndex == -1 : return [passing,0,0,False]
        if firstPass > specPeriod*0.90 and pf.find('...') > -1 : #fail within 10% of spec
            passRange = pf[pf.find('...'):pf.rfind('...')+2] #fails between big passes
            if periodIndex > pf.find('...') and periodIndex < pf.rfind('...')+2\
                 and '*' in passRange : holes = True
        # find pass right before the next fail
        nextFail = pf[firstIndex:].find('*') - 1
        nextFailPeriod = minPeriod + (firstIndex + nextFail)* periodStepSize
        # if the range between the first pass and the next fail contains spec period
        if nextFail > 0 and specPeriod <= nextFailPeriod and specPeriod >= firstPass :
            return [passing,round(firstPass,6),round(nextFailPeriod,6), holes]
        # find last two passes in a row
        lastIndex = pf.rfind('..') + 1
        lastPass = minPeriod + lastIndex * periodStepSize
        passingRange = pf[firstIndex:lastIndex]
        if '*' in passingRange: # pass-->fail-->pass
            holes = True
        return [passing, float(round(firstPass,6)),float(round(lastPass,6)), holes]
       
def examine_shmoo(filename, params, specPeriod, percentage, debugMode) :
    ''' Analyzes a file to find the minimum voltage that 
    still works for the given spec period
    Inputs:
        filename    - file to be read for parameters
        params      - parameters from the file header
        specPeriod  - given spec period for device
        percentage  - to calculate at Vnom +/- percentage
        debugMode   - boolean for printing debug statements
    Output:
        3 item list:
        (0) min/maxV - smallest or largest passing voltage 
        (1) vPercent - Vmin/Vmax aka. spec Voltage +/- %
        (2) notes   - notes about issues in test
    '''
    with open(filename,'r') as file :
        minPeriod = float(params[1])
        #maxPeriod = float(params[2])
        periodStepSize = float(params[3])
        minVoltage = float(params[4])
        maxVoltage = float(params[5])
        voltStepSize = float(params[6])
        voltage = params[7] # don't make float so we check how many sig figs
        flippedX = params[9]
        debugMode = int(debugMode)
        notes = '' 

        try:
            decimal = percentage
            if '%' in percentage :
                decimal = float(re.sub('[^0-9-]','',percentage))/100
            elif -1.00 <= float(percentage) <= 1.00:
                decimal = float(percentage)
            elif 1 < float(percentage) < 100 or -100 < float(percentage) < -1 :
                decimal = float(percentage)/100
            else: return 'Cannot read given percentage: ' + percentage
        except:
            return 'Cannot read given percentage: ' + percentage
        multiplier = 1 + decimal
        
        sigFigs = len(str(voltage)[str(voltage).find('.')+1:])
        voltage = float(voltage)
        vPercentRaw = round(multiplier*voltage, sigFigs) # exact spec Vmin/max number
        if vPercentRaw < minVoltage or vPercentRaw > maxVoltage: 
            percentage = str(decimal*100) + '%'
            errorString = '\n'+ percentage + '(' + str(vPercentRaw) + 'V) Not in Range ['+ \
                str(minVoltage) + 'V,'+ str(maxVoltage) + 'V]'
            return errorString
        vPercent = None # spec Vmin/max that actually exists in the graph
        # find Vnom + percentage% voltage
        if decimal < 0:
            vPercent = maxVoltage
            vAsInt = round(vPercent*1000) # larger numbers hold value better
            while vAsInt> round(1000*vPercentRaw):
                vAsInt = round(vAsInt - 1000*voltStepSize)
                vPercent = round(vAsInt/1000, sigFigs)
        else : 
            vPercent = minVoltage
            vAsInt = round(vPercent*1000)
            while vAsInt < round(1000*vPercentRaw):
                vAsInt = round(vAsInt + 1000*voltStepSize)
                vPercent = round(vAsInt/1000, sigFigs) 
        if vPercent < minVoltage or vPercent > maxVoltage: 
            percentage = str(decimal*100) + '%'
            errorString = '\n'+ percentage + ' (' + str(vPercent) + 'V) Not in Range ['+ \
                str(minVoltage) + 'V,'+ str(maxVoltage) + 'V]'
            return errorString
        
        # find the minimum/maximum working voltage
        contents = file.readlines()
        plotFound = False
        lowestV = float(999) 
        maxV = float(-999)       
        holes = []; fails =[]; failCtr = 0; failed = False
        pRange = [0,0] # period passing range at Spec voltage
        for line in contents :
            passFailStart = line.find('|')
            if failCtr > 2 : failed = True
            if failed : passFailStart = -1
            if passFailStart >= 0 : # lines that contain data points
                plotFound = True
                if debugMode > 1: print(line.strip('\n')) #print plot for debug
                lineVoltage = float(line[ :passFailStart])
                passing, pMin, pMax, holy = \
                    passing_range(line, periodStepSize, minPeriod, specPeriod, flippedX)
                if not passing  : # no passing range
                    if maxV != -999 and maxV < voltage*1.05: failCtr +=1
                    if lineVoltage == vPercent :
                        pRange = [pMin,pMax]
                        notes = 'Spec period does not pass at %s; '% lineVoltage
                        failed = True
                    if decimal < 0 and lowestV <= vPercent : # if the lowestV is already to spec
                        failed = True
                    else : fails.append(str(lineVoltage)) 
                    continue
                failCtr = 0
                if lineVoltage == vPercent :      
                    pRange = [pMin,pMax]
                    if specPeriod == pMin or specPeriod == pMax :
                        notes += 'Period on edge of passing period range; '
                # check if voltage works with spec period
                if specPeriod >= pMin and specPeriod <= pMax : 
                    # if there are holes but lowestV is already to spec, break
                    # if decimal < 0 and lowestV <= vPercent and holy: 
                    #     failed = True
                    #     continue
                    lowestV = lineVoltage
                    if maxV == -999 : # first working voltage
                        maxV = lineVoltage 
                else : fails.append(str(lineVoltage))
                if holy : # check line for holes
                    holes.append(str(lineVoltage))
            elif plotFound and debugMode > 1: # print bottom of plot for debug
                if len(line) > 1 : 
                    print(line.strip('\n'))
        if plotFound == False :
            raise ValueError('Shmoo Plot was Not Found for ' + filename)
        if lowestV == 999 or lowestV > voltage or float(maxV) < voltage:
            notes = 'All Failed'
        else:
            if len(fails) > 0 :
                Fails = []
                for fail in fails :
                    if float(fail) > float(lowestV) and float(fail) < float(maxV) :
                        Fails.append(str(fail))
                if len(Fails) > 0 :notes = 'Failing at: ' + '; '.join(Fails) + '; '
            if len(holes) > 0:
                Holes =[]
                for hole in holes :
                    if float(hole) > float(lowestV) and float(hole) < float(maxV) :
                       Holes.append(hole) 
                if len(Holes) > 0 : notes = notes + 'Holes at V: ' + '; '.join(Holes)
    if lowestV < voltage and decimal > 0 : # lookig for Vmax
        return [str(maxV), str(vPercentRaw), str(notes), pRange]
    if maxV > voltage and decimal < 0 :  # looking for Vmin   
        return [str(lowestV), str(vPercentRaw), str(notes), pRange]
    else : return [None, str(vPercentRaw), 'All Failed', pRange]

def print_params(parameters, device, temp) :
    ''' Prints out the parameters found in the given file'''
    print('\n        Test: ', (device + ', '+ temp).strip('\n'))
    print('Pattern Name: ', parameters[0].strip('\n'))
    print('        Vnom: ', str(parameters[7]).strip('\n'), 'V')
    print('    EQNSET #: ', parameters[8].strip('\n'))
    print('      Domain: ', parameters[1].strip('\n'), ' - ',\
        parameters[2].strip('\n'), 'ns')
    print('       Range: ', parameters[4].strip('\n'), ' - ',\
        parameters[5].strip('\n'), 'V')
    print(' X Step Size: ', parameters[3].strip('\n'))
    print(' Y Step Size: ', parameters[6].strip('\n'))
    print('    flippedX: ', parameters[9])

def more_recent(testA,testB) :
    ''' Finds the most recently created test based off of time in the name and
    the date of creation
    Input:
        Two test CSV lines with same test parameters 
    Output:
        Most recent test lines. If same date then first test. If error both tests
    '''
    tsA = testA[testA.rfind(','):].split('_')
    tsB = testB[testB.rfind(','):].split('_')
    try :
        for i in range(0,len(tsA)):
            tsA[i] = re.sub('[^0-9]','',tsA[i])
            tsB[i] = re.sub('[^0-9]','',tsB[i])
            if tsA[i] > tsB[i] :
                return testA
            elif tsA[i] > tsB[i] :
                return testB
    except : return testA + testB
    return testA # return only one if they are equal

def info_from_pathname(fullPath) :
    device = temp = None
    pCorners = ['FF','SS','TT','FS','SF']  
    parents = fullPath.split(os.sep)
    #get device name
    for parent in parents :
        for pc in pCorners:
            deviceInd = parent.find(pc)
            try:
                if pc in parent and '_' in parent[deviceInd:] : #goes until _ found
                    num = re.sub('[^0-9]','',parent[:parent[deviceInd:].find('_')+deviceInd])
                    if len(num) > 0 and num.isdigit() : device = pc + num  
                elif pc in parent :
                    num = ''
                    afterDev = parent[deviceInd+2:]
                    for char in afterDev : #check all chars after the PCorner
                        if char.isdigit() :
                            num = num + char
                        else : break # stops after it finds a non-digit
                        if len(num) > 0 and num.isdigit() : device = pc + num
            except: pass
    # get temperature
    for parent in parents :
        untilTemp = parent[:parent.find('C')]
        try:
            if 'C' in parent: 
                if '_' in untilTemp : #from _ on if it eists
                    num = re.sub('[^0-9-nm]','',untilTemp[untilTemp.find('_'):])
                else : # get all nums before 'C'
                    num = re.sub('[^0-9-nm]','',untilTemp)
                num = re.sub('[nm]','-',num) # negative sign
                if '-' in num : num =  '-' + re.sub('[^0-9]','',num)
                try: 
                    num = int(num)
                    if num > -100 and num < 1000 : #if temp between (-100,1000)
                        temp = str(num) + 'C' 
                except: pass
        except: pass
    return [temp,device]
    
def process_shmoo_text(inputFiles,outputFile,eqnset,spec,device,temp,percentage,debug) :
    ''' Controller module for analyzing shmoo files, individually or in a list.
    Converts one shmoo file and and finds the lowest voltage where it is still 
    passing at spec period 
    Input(command line arguments):
        input       --> list of file/files to be read
        outputFile  --> .csv name of file to write results to
        eqnset      --> file to find spec period in
        spec        --> spec period for test
        device      --> device name
        temp        --> testing temperature
        percentage  --> to calculate at Vnom +/- percentage
        debug       --> turn on debugging print statements (mode 0-3)
    Output: 
        Data for one shmoo plot, to .csv file if given or to terminal if not 
    '''
    # check if input and output files are valid
    for file in inputFiles :
        if not os.path.isfile(file) :
            return print('%s is not a file' % file) 
        if not(file.endswith('.txt')) :
            return print('\nPlease provide .txt inputs')

    # check output directory
    if outputFile != None :
        outputDir = os.path.abspath(os.path.join(outputFile, os.pardir))
    else: outputDir = '.'
    if outputFile != None and not (outputFile.endswith('.csv')) :
        return print('\nPlease provide a .csv output file name')
    elif outputFile != None and not os.path.isdir(outputDir) : 
        os.makedirs(outputDir)
    if not os.access(outputDir, os.W_OK) : 
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
        percentageOut = ('%f' % (decimal*100)).rstrip('0').rstrip('.') + '%'
    except: pass
    # examine the shmoo for every input file        
    for file in inputFiles :
        # get temp and device if not given
        if device == None or temp == None :
            fullPath = os.path.abspath(file)
            getInfo = info_from_pathname(fullPath)
            if temp == None : temp = getInfo[0]
            if device == None : device = getInfo[1]
        if temp == None or device == None:
            if debug > 0 and __name__ != '__main__' :
                return 'Cannot find Temp and Device for test folder' + input
            if temp == None : 
                print('Please provide temperature (-T temp)')
            if device == None : 
                print('Please provide device name (-D name)')
            return
        parameters = params(file,outputDir)
        if isinstance(parameters,str): 
            if __name__ != '__main__':
                return parameters
            else: print(parameters)
        
        #check if all parameters are numbers besides the pattern name 
        for x in range(1,8) :
            try: 
                float(parameters[x])
            except ValueError:
                print('\nCannot read input file: ', file)
                if debug > 0 :
                    print_params(parameters,device,temp)
                return '' 
        # find the spec periods, either from file or from given number
        if spec == None : specPeriodsFinal = []
        else: specPeriodsFinal = [spec]
        if spec == None and eqnset == '.' : # if neither are provided, check cwd 
            if os.path.isfile('./timingSpecs.txt') :
                eqnset = './timingSpecs.txt'
            elif os.path.isfile(os.path.join(outputDir,'./timingSpecs.txt')) :
                eqnset = os.path.join(outputDir,'./timingSpecs.txt')#if in output dir
            else :
                return print('Please provide EQNSET file (-eq) or spec Period (-s)')
        if not eqnset == '.' :
            name = os.path.relpath(os.path.realpath(eqnset)) 
            with open(eqnset, 'r') as eq :
                lines = eq.readlines()
            try :
                # find eqnset #N period
                periodLines = []
                for line in lines :
                    eqnStr = 'EQNSET' + parameters[8] + ','
                    if line.find(eqnStr) >= 0:
                        periodLines.append(line)
                    if 'TIMINGSET' in line :
                        patName = line[line.find('\"')+1:line.rfind('\"')] 
                        if parameters[0].lower().find(patName.lower()) >= 0 : #[:len(patName)] after parameters[0]
                            periodLines.append(line)
                        
                # gets list of possible periods, must be in comma separated list
                specPeriodsFinal = []
                for periodLine in periodLines :
                    values = periodLine.split(',')[1:]
                    for value in values:
                        periodStr = str(re.sub('[^0-9.]', '', value))
                        try :
                            float(periodStr)
                        except: continue
                        if len(periodStr) == 0:
                            continue
                        #if not 'EQNSET' in specPeriods[period-1] : fromName = True
                        #else: fromName = False
                        #if len(specPeriodsFinal) > 0 and fromName : continue
                        if float(periodStr) < float(parameters[1]) or \
                            float(periodStr) > float(parameters[2]) : # if not in boundaries 
                            if debug > 1 : 
                                print('\nSpec period ' + str(periodStr) + ' not in boundaries for ' + \
                                    parameters[0] + ', ' + device + ' ' + temp, end='')
                            continue
                        specPeriodsFinal.append(periodStr)
                if len(specPeriodsFinal) == 0 : # if no working spec Periods 
                    raise Exception
            except Exception as e:
                if debug > 0 :
                    if __name__ == '__main__' :
                        print('\nNo spec period for ' + parameters[0] + ' in ' +\
                                name + ', EQNSET ' + parameters[8] + ' (' + parameters[1] \
                                    +'ns-' + parameters[2] + 'ns)', end='') 
                return '\nNo spec period for ' + parameters[0] + ' in ' + name +\
                        ', EQNSET ' + parameters[8] + ' (' + parameters[1] +'ns-' + \
                            parameters[2] + 'ns)'

        parent = os.path.abspath(os.path.join(file, os.pardir))
        pName = os.path.basename(os.path.normpath(parent))
        # get pattern name and issues 
        if 'shmoo' in pName or 'Shmoo' in pName or 'shm' in pName:
            parameters[0] = pName
        patternName = parameters[0] 
        
        # analyze the plot with all the spec periods from the timing doc
        for sP in specPeriodsFinal :
            results = ''
            try:
                results = examine_shmoo(file, parameters, float(sP), percentage, 0)
                if isinstance(results,str) : print(results); raise AttributeError # if error in percentage
                printing = False
                vActual, vSpec, notes, pRangeAtSpec = results
                # print graphs that haves holes or failing lines
                if debug > 1 and ('Failing at' in notes or 'Holes at' in notes) :
                    printing = True
                #print graphs that have spec periods on the edge of the passing range
                elif debug > 2 and ('Period on edge' in notes or 'All Failed' in notes) :
                    printing = True
                #prints all graphs and params
                elif debug > 3 :
                    printing = True
                if printing:
                    print_params(parameters, device, temp)
                    if '-' in percentage: 
                        print('        Vmin:  %s (Vnom%s)' %(vSpec,percentageOut))
                    else :
                        print('        Vmax:  %s (Vnom+%s)' %(vSpec,percentageOut))
                    print(' Spec Period: ',sP, '\n       Notes: ', notes.strip(';'))
                    examine_shmoo(file, parameters, float(sP), percentage, debug)
            except ValueError as ve:
                print(str(ve))
                return str(ve)
            except AttributeError : # percentage not able to be analyzed
                if __name__ == '__main__' :
                    print(results, end=''); 
                    print_params(parameters,device,temp)
                else: 
                    return results + ' for: ' + parameters[0]
            except Exception as e :
                if debug > 0 :
                    print('\nNo results in %s for period %s, %s %s' % (pName, sP, \
                    device, temp), end='')
                    print_params(parameters, device, temp)
                    print(' Spec Period: ', sP)
                    if type(e) == ValueError : 
                        print(str(e))
                    else :
                        examine_shmoo(file, parameters, float(sP), percentage, debug)
                continue
            if vActual == None :
                fail = 'FAIL'
                holes =''
                badV =''
            else :
                if 'Holes' in notes: holes = 'VH'
                else : holes = ''
                if vActual > vSpec and '-' in percentage: badV = 'Vmin'
                elif vActual < vSpec and not '-' in percentage: badV = 'Vmax'
                else : badV = ''
                if str(vSpec) in notes : fail = str(vSpec)
                else : fail = ''
            if '-' in percentage : pActual = pRangeAtSpec[0]
            else : pActual = pRangeAtSpec[1]
            if pRangeAtSpec == [0,0] :
                pActual = 'FAIL'
                tFail = 'FAIL'
                tHoles =''
                badT =''
                tNotes = 'All periods fail @ spec Voltage'
            else :
                tFail = ''
                tHoles = ''
                tNotes = ''
                if 'Spec period does not pass ' in notes :
                    tNotes = 'Spec period does not pass at Vmin'
                if 'Period on edge' in notes :
                    tNotes += 'Period on edge of passing range;'
                if 'Holes' in notes :
                    holesList = notes[notes.find('Holes'):]
                    if str(vSpec) in holesList : 
                        tHoles = 'TH'
                        tNotes += 'Holes at V: '+vSpec+';'
                if  float(pActual) > float(sP) and '-' in percentage: badT = 'Tmin'
                elif float(pActual) < float(sP) and not '-' in percentage: badT = 'Tmax'
                else : badT = ''
            if 'Period on edge of passing period range; ' in notes :# only for T
                notes = notes.replace('Period on edge of passing period range; ','')
            #get timestamp of file 
            created = datetime.fromtimestamp(os.path.getctime(file))
            try:
                fileName = os.path.basename(file)
                msIndex = fileName.rfind('ms')
                # ms = int(fileName[fileName[:msIndex].rfind('s') +1 : msIndex])
                fileName = fileName[:msIndex]
                sIndex = fileName.rfind('s')
                s = int(fileName[fileName[:sIndex].rfind('m') +1 : sIndex])
                fileName = fileName[:sIndex]
                mIndex = fileName.rfind('m')
                m = int(fileName[fileName[:mIndex].rfind('h') +1 : mIndex])
                fileName = fileName[:mIndex]
                hIndex = fileName.rfind('h')
                h = int(fileName[fileName[:hIndex].rfind('_') +1 : hIndex])
                fileName = fileName[:hIndex]
                dayIndex = fileName.rfind('_')
                day = int(fileName[fileName[:dayIndex].rfind('_') +1 : dayIndex])
                fileName = fileName[:dayIndex]
                month = int(created.strftime('%M'))
                year = created.year
                timeStamp = '_'.join([str(num) for num in [year,month,day,h,m,s]])
            except:
                timeStamp = '_'.join([str(num) for num in [created.year, \
                    created.month,created.day,created.hour,created.minute,\
                        created.second]])

            vOutput= '%s, %s, %s, %s, %s, %sns, %s, %s, %s, %s, %s, %s, %s, %s, %s\n' % \
                (patternName, parameters[8], holes, badV, fail, sP, parameters[4], \
                    parameters[5], vSpec, temp, device, vActual, notes,\
                        timeStamp, file)
            tOutput= '%s, %s, %s, %s, %s, %s, %s, %s, %sns, %s, %s, %s, %s, %s, %s\n' % \
                (patternName, parameters[8], tHoles, badT, tFail, vSpec, parameters[1], \
                    parameters[2], sP, temp, device, pActual, tNotes,\
                        timeStamp, file)
            if '-' in percentage :
                CSVHeader = 'Pattern Name, EQNSET #, Holes, Bad Vmin, Uncertain Vmin,'\
                + ' Spec Period(ns), From (V), To (V), Spec Vmin(%s), Temperature,'% percentageOut\
                + ' Device, Recorded Vmin (V), Notes, Date, File Path \n'
            else :
                CSVHeader = 'Pattern Name, EQNSET #, Holes, Bad Vmax, Uncertain Vmax,'\
                + ' Spec Period(ns), From (V), To (V), Spec Vmin(+%s), Temperature,' % percentageOut\
                + ' Device, Recorded Vmax (V), Notes, Date, File Path \n' 
            TCSVHeader= CSVHeader.replace('Vm','Tm').replace('(V)','(ns)')
            TCSVHeader = TCSVHeader.replace('Spec Period(ns)','Spec Vmin(%s)'% percentageOut)
            TCSVHeader = TCSVHeader.replace('Spec Tmin(%s)'% percentageOut,'Spec Period(ns)')
            if outputFile == None or printing :
                # vOutput = vOutput.replace(', '+file,'')
                # tOutput = tOutput.replace(', '+file,'')
                print('CSV outputs: ')
                if '-' in percentage:
                    print('Vmin:\n'+vOutput.replace(';',','))
                    print('Tmin:\n'+tOutput.replace(';',','))
                else: 
                    print('Vmax:\n'+vOutput.replace(';',','))
                    print('Tmax:\n'+tOutput.replace(';',','))
                print('-'*100)
            if outputFile :  #write the line to the csv and replace it if it's already there
                #voltage CSV
                lines = []
                written = False
                if os.path.isfile(outputFile) : # if output exists already
                    with open(outputFile, 'r+') as oldFile :
                        content = oldFile.read()
                        info = '%sns, %s, %s, %s, %s, %s' % \
                            (sP, parameters[4], parameters[5], vSpec, temp, device)
                        if (content.find((patternName+', '+ parameters[8])) < 0) or\
                            (content.find(info) < 0): # if test not in file
                            oldFile.write(vOutput); written = True
                    if written == False : 
                        with open(outputFile, 'r') as oldFile :   
                            lines = sorted(oldFile.readlines()) 
                if written == False : 
                    with open(outputFile,'w') as newFile :
                        newFile.write(CSVHeader) #write header at top
                        repeat = ''
                        for line in lines: 
                            # check if there is already something written for the file 
                            if 'Pattern Name' in line : continue # skip header
                            lineItems = line.split(', ')
                            if (patternName in lineItems) and (temp in lineItems) and \
                                (device in lineItems) and (str(vSpec) in lineItems)\
                                    and ((sP + 'ns') in lineItems) and \
                                        str(parameters[8]) in line:
                                        repeat = line # for checking most recent
                            else : newFile.write(line) # copy the unchanged lines
                        newFile.write(more_recent(vOutput,repeat))
                #period CSV
                lines = []
                tOutputFile = os.path.join(os.path.dirname(outputFile),\
                    os.path.basename(outputFile).replace('Vm','Tm'))
                if os.path.isfile(tOutputFile) : # if output exists already
                    with open(tOutputFile, 'r+') as oldFile :
                        content = oldFile.read()
                        info = '%s, %s, %s, %sns, %s, %s' % \
                            (vSpec, parameters[1], parameters[2], sP, temp, device)
                        if (content.find((patternName+', '+ parameters[8])) < 0) or\
                            (content.find(info) < 0): # if test not in file
                            oldFile.write(tOutput); break
                    
                    with open(tOutputFile, 'r') as oldFile :   
                        lines = sorted(oldFile.readlines()) 
                with open(tOutputFile,'w') as newFile :
                    newFile.write(TCSVHeader) #write header at top
                    repeat = ''
                    for line in lines: 
                        # check if there is already something written for the file 
                        if 'Pattern Name' in line : continue # skip header
                        lineItems = line.split(', ')
                        if (patternName in lineItems) and (temp in lineItems) and \
                            (device in lineItems) and (str(vSpec) in lineItems)\
                                and ((sP + 'ns') in lineItems) and \
                                    str(parameters[8]) in line:
                                    repeat = line # for checking most recent
                        else : newFile.write(line) # copy the unchanged lines
                    newFile.write(more_recent(tOutput,repeat))

if __name__ == '__main__':
    for arg in range(0,len(sys.argv)) :
        if '%' in sys.argv[arg] : # in case the percentage sign messes up argparse
            sys.argv[arg] = sys.argv[arg].replace('%','') 
    parser = argparse.ArgumentParser(description='Analyze a shmoo plot files', \
        formatter_class = argparse.RawTextHelpFormatter,\
        epilog = 'usage examples:\n'\
        '  process_shmoo_text -i shmooText/Bscan_1p8_shmoo/PartId_P1.1_Site_1_Mode_'\
        'PassFail_Apr_26_10h43m00s111ms.txt -eq ../char/timingSpecs.txt -d 4\n\n'\
        '  process_shmoo_text -i snTT2_25C/shmooText/Temp_sensor_Burst_VDD08_'\
        'Shmoo/PartId_P1.1_Site_1_Mode_PassFail_Apr_26_11h37m59s754ms.txt'\
        '-d 2 -p -7% -D TT2 -T 25C\n')
    parser.add_argument('-v', '-V', '--version', dest='version', action='store_true',\
        default=False, help='get version of script and exit')
    parser.add_argument('-i', '--input', dest='inputFiles', nargs='+', \
        default=None, help='required input file names (.txt only) single or list ')
    parser.add_argument('-o', '--output', dest='outputFile', default=None, \
        help='output file path(.csv only w/ Vmin/Vmax in name). DEFAULT print to terminal')
    parser.add_argument('-eq', '--timing', dest='eqnset', default='.', \
        help='timing specs file location; default checks if in current folder.\n' +\
            'can also give *.tim file. (use "get_timing" to get timing spec file)')
    parser.add_argument('-s', '--spec', dest='spec', default=None, \
        help='spec period for listed files')
    parser.add_argument('-D', '--device', dest='device', default=None, \
        help='device name (eg. TT02)')
    parser.add_argument('-T', '--temperature', dest='temp', default=None, \
        help='test temperature (eg. 25C)')
    parser.add_argument('-p', '--percent', dest='percentage', default='-5%%', \
        help='percentage from Vnom(+/-). DEFAULT = -5%%')
    parser.add_argument('-d', '--debug', dest='debug', type=int, default=0, \
        choices=range(0,5), metavar='[0-4]', help='debug mode [0-4]\n' \
            '0 = No Messages\n' '1 = Error Mesages\n' '2 = Some Bad Plots\n' \
                '3 = All bad Plots w/ Notes\n' '4 = Print All Files Listed')
    args = parser.parse_args()
    if args.version :
        version = get_version(sys.argv[0])
        print('Version: ' + version)
        sys.exit()
    if args.inputFiles == None :
        parser.print_usage()
        print('Input files are required. Please include using -i')
        sys.exit()
    process_shmoo_text(args.inputFiles, args.outputFile, args.eqnset, args.spec,\
         args.device, args.temp, args.percentage, args.debug)