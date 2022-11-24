#import relevant libraries

import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os
import statistics
import xlsxwriter



#This module imports data from the excel file as a list
def extractfromexcel(filename):
    data1= pd.read_excel(filename) 
    data2= data1.to_numpy()  #imports rows from spreadsheet
    data3 = data2.tolist() #converts dataframe to list for easier handling. Data3 is a list of lists where ever row in the excel sheet is a list.
    return data3



    
#This module combines and holds information for a given tutor group into one array (helddata)
def mergetutorrows(checkingdata, helddata):

    for i in range(0,len(helddata),1):  #this should work because tutors should have the same number of students
        checkingdatarow = checkingdata[i]    #remove name from checking data since we already heave it from first pass
        del checkingdatarow[0]
        a = helddata[i]
        helddata[i] = a + checkingdatarow #combine register entries strings for students who should appear in same order
    return helddata




#This module extracts data into storage arrays
def extractinfomation(currenttutorgroup, helddata, processeddata):
    tutorinfo = ['Scope:', currenttutorgroup]
    helddata.insert(0,tutorinfo) #insert tutor information at start of group data
    processeddata = processeddata + helddata
    return processeddata




#This module processes extracted data so that all the information of each tutor group is combined together into single rows for further analysis

def processdata(data3):

    #create empty arrays for storage
    processeddata = []
    helddata = []
    checkingdata = []
    currenttutorgroup = []
    counter = 0

    #document pre-processing (matching multiple attendance blocks to the same student)
    for j in range(0,len(data3),1): #for each row of the excel document
        
        row = data3[j];
        
        counter = counter + 1
        
        if isinstance(row[0],str)==True: #only look at columns that start containing a string 

            if ('Official Register' in row[0])== False and ('Period:' in row[0])== False and ('Missing' in row[0])== False: #if cell is a string (name) but not a Header

                if ('Scope:' in row[0]) == True: #we need to extract this information to create the tutor group folders
                    if len(currenttutorgroup) < 1: #if its the first time we come across a tutor group then simply extract
                        currenttutorgroup = row[1]
                    else:
                        if currenttutorgroup == row[1]: #if we have already passed a given tutor group label
                            if len(helddata)<1: # if no attendance data have been passed yet
                                helddata = checkingdata  #checked data passed onto held data
                                checkingdata = [] #resetcheckingdata

                            else: 
                                #if tutor group is still the same and data is already being held then combine and hold dataset
                                helddata = mergetutorrows(checkingdata,helddata)
                                checkingdata = [] # reset data check
                            
                                    
                        else: #if tutor group is for a new class

                            #Update held data from last round
                            helddata = mergetutorrows(checkingdata,helddata)

                            #Extract data and reset storage arrays
                            processeddata = extractinfomation(currenttutorgroup, helddata, processeddata)
                            helddata = []
                            checkingdata = []
                            currenttutorgroup = row[1] #update tutor group

                else: #if its a student name and their register entries then just store into temporary array of data being checked
                    checkingdata = checkingdata + [row]

        #if reached final row then there wont be a tutor label to flag data extraction - so set it to happen
        if j == len(data3)-1:
            
            #Pass on held data from last round
            if len(helddata)<1: # if no elements have been passed yet
                helddata = checkingdata  #checked data passed onto held data
                checkingdata = [] #resetcheckingdata

            else:
                #combine and hold data
                helddata = mergetutorrows(checkingdata,helddata)

            #Extract data
            processeddata = extractinfomation(currenttutorgroup, helddata, processeddata)
                    
        print("Processing data " + str(counter) + " out of " + str(len(data3)))

    return processeddata





    
#decimal place formatters
my_formatter="{0:.0f}"  #zero decimal places
my_formatter1="{0:.1f}" #one decimal place
my_formatter2="{0:.2f}" #two decimal places



# import data from excel file
filename = "data.xlsx"  #update name of file to be opened
data3 = extractfromexcel(filename)


# process data
processeddata = processdata(data3)
    


#Compute attendance data

for row in processeddata:
    
    studentinfo = [] #reset studentinfo for each student so we can save each graph in turn

    if 'Scope:' in row[0]: #if the row starts with scope then we know we have reached a new tutor group and we need to create a new folder

        #Create new directory for tutor group
            
        #extract tutor group infor from cells
        tutorstring = row[1]
        end = len(tutorstring)
        tutorcode = tutorstring[end-2:end]
        if tutorcode[0]=='0':   #if tutor code starts with a '0' for year 10 then add the 1 to the code
            tutorcode = '1' + tutorcode
            
        #Create a new folder
        currentdir = os.getcwd() # get current directory
        foldername = tutorcode #define folder name
        path = os.path.join(currentdir, foldername) #create a path string
        os.mkdir(path) # create directory
            

    else: #if its not the tutor row its a student data row
            
        daycounter = 0
        attended = 0
        week = 0
        notenrolledweek = 0
        flaggedZ = 0
        relevantinfo = []
        nanflag = 0

            
        for i in range(0,len(row),1):  #for each register mark and empty space in the student data row
    
            entry = row[i]
                
            if i == 0: #if its the row element containing the name
                relevantinfo.append(entry) # collect name
                
            elif isinstance(entry, str)==True: #if its not an empty nan column then compute the register mark

                if entry == '/\\': #if present all day
                    daycounter = daycounter + 1 #acknowlegde school was open
                    attended = attended + 1  #acknowledge child was in

                elif entry != '/\\' and ('/' in entry or '\\' in entry): #if present for half a day eg. for unauthorised mornign/afternoon appointments or Us)
                    if ('C' not in entry) and ('L' not in entry) and ('P' not in entry) and ('V' not in entry) and ('W' not in entry) and ('R' not in entry): #these codes should not penalise attendance
                        daycounter = daycounter + 1 #acknowlegde school was open
                        attended = attended + 0.5 # acknowlegde child was in for half a day

                    else:
                        daycounter = daycounter + 1 
                        attended = attended + 1
                        
                elif ('CC' in entry) or ('LL' in entry) or ('PP' in entry) or ('VV' in entry) or ('WW' in entry) or ('RR' in entry):  #for all codes of authorised absence
                    daycounter = daycounter + 1 
                    attended = attended + 1
                        
                elif '#' in entry:
                    continue
                    
                elif'Z' in entry: #if child appears in system but hasnt been enrolled yet then flag it up
                    flaggedZ= flaggedZ +1
                    
                else: #for any other codes (eg. M, O or I) just aknowledge the day but dont mark attendance
                    daycounter = daycounter + 1 #acknowlegde school was open

            else: #if its an empty nan entry

                previousentry = row[i-1] #check if there was register data in the previos row entry
                    
                if isinstance(previousentry, str) == True: #if this is the first nan column reached between weekly data
                    
                    if daycounter == 0 and flaggedZ ==0: #if no Z codes have been flagged and school was not open then reset counters and skip week
                        daycounter = 0 #reset counters for next set of columns of weekly attendance
                        attended = 0

                    elif daycounter == 0 and flaggedZ > 0: #if student wasnt on roll yet
                        notenrolledweek = notenrolledweek + 1 #track weeks to then shift chronology along so their arrival matches the correct week
                        flaggedZ = 0  #reset flag
                            
                    else: #if school was open during the week and the student was not marked as a Z then calculate percentage session attendance for that week
                        weekattendance = (attended/daycounter)*100
                        relevantinfo.append(weekattendance) #extract weekly percentage attendance for each week-set of columns
                        week = week +1 #update how many weeks have been attended
                        daycounter = 0 #reset counters for next set of columns of weekly attendance
                        attended = 0


            #if loop has reached the final row entry and final entry is not a nan then compute the final count since there is no nan to flag the end of the week has been reached
            if i == len(row)-1:
                if isinstance(row[i], str) == True:
                    weekattendance = (attended/daycounter)*100
                    relevantinfo.append(weekattendance)
                    daycounter = 0 #reset counters for next set of columns of weekly attendance
                    attended = 0
                    week = week +1
                    

        studentinfo = studentinfo + relevantinfo #extract student data before restarting the loop and moving on to the the next row
        

        #Before analysing data in next student row - plot and save in newly created folder

        name = studentinfo[0] #student name is first entry and percentage attendance are the remaining list entries
        end = len(studentinfo)
        studentdata = studentinfo[1:end]
               
        #plot and save graph

        #Create x axis for time course graph - this depends on when student became enrolled
        weekssofar  = notenrolledweek + week
        if notenrolledweek >0:
            weeks = np.linspace(notenrolledweek,weekssofar-1,week) #create a vector with week numbers going from enrollment to however many week data points there are
        else:
            weeks = np.linspace(1,week,week)
        

        #create curve for linear regression where m1 is gradient and n1 is intersect
        m1, n1= np.polyfit(weeks, studentdata, 1) #function to fit polinomialy a curve with low degree of fitting (1)
        y1= m1*weeks+n1   #line of linear regression


        plt.subplots() #creates plot
        plt.scatter(weeks, studentdata, color="blue") #plots scatter plot of datapoints
        plt.plot(weeks,y1, "r-", label= "fit") #plots linear regression

        plt.xlabel(" Week ") #sets x label
        plt.xticks(weeks)
        plt.ylabel("Attendance (%)")# sets y label
        plt.ylim([0, 100])
        plt.grid() #applies a grid behind

        graphname = name #graph will have the name of the student
        saveto = path + '/' + name + '.png'
            
        plt.savefig(saveto, dpi = 200) #with with specific dpi resolution
        plt.close()  #close figure to save memory
        print(name)




