#Uzair Jawaid
#Python script to automate monthly incident trending reports
#2019-06-01

#Import necessary libraries for excel support
import pandas as pd
import time
import matplotlib.pyplot as plt

#Set start time to calculate run time of script
start_time = time.process_time()

#Read the excel files
file1 = pd.read_excel (r'C:\Users\Uzair\Desktop\Work\Automation\file1.xlsx')
print ("Reading first file...")

file2 = pd.read_excel (r'C:\Users\Uzair\Desktop\Work\Automation\file2.xlsx')
print ("Reading second file...")

file3 = pd.read_excel (r'C:\Users\Uzair\Desktop\Work\Automation\file3.xlsx')
print ("Reading third file...")

#Create a DataFrame to hold all the files data
all_data = pd.DataFrame()
all_data = all_data.append(file1, ignore_index = True) #append the sheet to the DataFrame
print ("Appending files...(1/3)")
all_data = all_data.append(file2, ignore_index = True)
print ("Appending files...(2/3)")
all_data = all_data.append(file3, ignore_index = True)
print ("Appending files...(3/3)")

print ("\n", "Combined data: ")
print("\n", all_data) #Display all data

all_data.to_excel(r'C:\Users\Uzair\Desktop\Work\Automation\mergedFile.xlsx') #Save DataFrame as a new Excel file

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

print("\n", errorDB)
errorDBLabels = ["Error Type", "Occurrences"] #labels for the new data frame
errorData = pd.DataFrame.from_records(errorDB, columns = errorDBLabels) #data frame to display errors vs occurances
print("\n", errorData)

errorData.to_excel(r'C:\Users\Uzair\Desktop\Work\Automation\analysisFile.xlsx') #Save DataFrame as a new Excel file

barGraph = pd.DataFrame({'Error': errorList, 'Occurances': errorCount})

barGraph.plot.bar(x='Error', y='Occurances', rot=0)
#plt.show()    #Currently the error names are too long to display nicely on the graph
  
#Show execution time for process
print ("\n", "Run-time: ", time.process_time() - start_time, "seconds to execute") 