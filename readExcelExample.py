#Uzair Jawaid
#Python script to automate monthly incident trending reports
#2019-06-01

#Import necessary libraries for excel support
import pandas as pd
import time

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

print ("Combined data: ")
print(all_data) #Display all data

#all_data.to_excel(r'C:\Users\Uzair\Desktop\Work\Automation\mergedFile.xlsx') #Save DataFrame as a new Excel file

print(all_data.shape) #Get the rows, columns
print(all_data["Heading 4"]) #Show only column 4, along with the data type and name of the column

errors = all_data["Heading 4"].count() #Get the number of rows in the last column
print("Number of errors: ", errors)

#Show execution time for process
print (time.process_time() - start_time, "seconds to execute") 