#Uzair Jawaid
#Python script to automate monthly incident trending reports
#2019-06-01

#Import necessary libraries
import pandas as pd
from pandas import ExcelFile
import time
import matplotlib.pyplot as plt

#Set start time to calculate run time of script
start_time = time.process_time()

#Read the excel files
file1 = pd.read_excel (r'week1.xlsx')
print ("Reading first file...")

file2 = pd.read_excel (r'week2.xlsx')
print ("Reading second file...")

file3 = pd.read_excel (r'week3.xlsx')
print ("Reading third file...")

file4 = pd.read_excel (r'week4.xlsx')
print ("Reading fourth file...")

#Create a DataFrame to hold all the files data
all_data = pd.DataFrame()
all_data = all_data.append(file1, ignore_index = True) #append the sheet to the DataFrame
print ("Appending files...(1/4)")
all_data = all_data.append(file2, ignore_index = True)
print ("Appending files...(2/4)")
all_data = all_data.append(file3, ignore_index = True)
print ("Appending files...(3/4)")
all_data = all_data.append(file4, ignore_index = True)
print ("Appending files...(4/4)")

print ("\n", "Combined data: ")
print("\n", all_data) #Display all data

all_data.to_excel(r'mergedFile.xlsx') #Save DataFrame as a new Excel file

print("\n", "Rows, Columns: ", all_data.shape) #Get the rows, columns
print("\n", all_data["MessageAlert"]) #Show only column 4, along with the data type and name of the column

errors = all_data["MessageAlert"].count() #Get the number of rows in the last column
print("\n", "Number of errors: ", errors)

errorList = [] #list of unique errors

#loop through rows
for x in range(errors):
        if all_data.iloc[x,5] not in errorList: #if the error isn't already defined in the list, add it to the list
                errorList.append(all_data.iloc[x,5])
                continue

print("\n", "Job errors query: ", errorList)

errorCount = [0]*len(errorList) #Empty list to hold the count of each error by index

#loop through the errors column
for y in range(errors):
        #loop through the unique errors column
        for z in range(len(errorList)):
                #check for a match, and add 1 to the count each time there's a repetition of the error in the main list
                if all_data.iloc[y,5] == errorList[z]:
                        errorCount[z] += 1
                else:
                        continue

print("\n", "Error occurances query: ", errorCount)

errorDB = [] #Create a series list to make a map of the error matching with its number of occurrences
#loop through the errors list
for a in range(len(errorList)):
        #add a tuple for each error into the errorDB
        errorDB.append((errorList[a], errorCount[a]))

incidentList = [] #2D Array to hold the format of [error[incidents]]
#Get incident list
for b in range(len(errorList)):
        smallerList = [] #smaller list is a grouping of all similar incidents by their INC# and AlertMessage
        incidentList.append(smallerList) #add a new list for each error
        for c in range(errors):
                if errorList[b] == all_data.iloc[c,5]: #check if there's a match
                        smallerList.append(all_data.iloc[c,0]) #insert the incident number into the list

#Inserting incident columns into new dataframe
incidentDF = pd.DataFrame()
for d in range(len(errorList)): #loop through error list
        for e in range(len(incidentList[d])):
                #set values to corresponding [error: incident] format
                #incidentDF.set_value(e, d, incidentList[d][e]) ---commented out because newer versions will remove this method---
                incidentDF.at[e,d] = incidentList[d][e]
#Setting column names to errors
incidentDF.columns = errorList

incidentDF.to_excel(r'incidentsList.xlsx')

print("\n", "New dataframe: ", incidentDF) #display incident dataframe
print("\n", errorDB)
errorDBLabels = ["Error Type", "Occurrences"] #labels for the new data frame
errorData = pd.DataFrame.from_records(errorDB, columns = errorDBLabels) #data frame to display errors vs occurances

#Calculating weekly avg occurances for each error
for f in range(len(errorList)):
        errorData.at[f,2] = errorCount[f]//4

errorDBLabels = ["Error Type", "Occurrences", "Weekly Average"] #labels for the new data frame
errorData.columns = errorDBLabels
print("\n", errorData)

errorData.to_excel(r'reoccuranceCount.xlsx') #Save DataFrame as a new Excel file

#create dataframe for bar graph
barGraph = pd.DataFrame({'Error': errorList, 'Occurances': errorCount})

barGraph.plot.bar(x='Error', y='Occurances', rot=0)
#plt.show()    #Currently the error names are too long to display nicely on the graph
  
#Show execution time for process
print ("\n", "Run-time: ", time.process_time() - start_time, "seconds to execute") 