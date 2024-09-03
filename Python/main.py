'''
--------------------------------------------------------------------
This is the test script for formulating the search problem of visiting 
Bucharest from Arad. The formulation includes defining the initial 
state, actions, successor function, goal test, path cost, and solution.

In additional, please visit the "Readme.md" file for more information
--------------------------------------------------------------------
'''

import GetCSVFile as gcv

if __name__ == "__main__":
    print("Setting up the environment...")

    # All set up goes here
    gcv.Checking.CheckingTheDefaultEnvironment("Inventory")

    print("Finish setting up. You can now use the program\n")

    print("Input \"end\" to end the program\n")

    # The main program start here
    while (True):

        # Get the command to manipulate the file
        Command = str(input("Input the command (\"Open\" or \"Save\"): "))

        # If the command is "Open", openning the file name
        if Command.lower() == "open":

            # Get the input file name
            FileName = str(input("Input the file name to display (you may include the file extension): "))

            if FileName.lower() == "end":
                break

            # Openning the file
            FILE = gcv.FileAttribute.OpenTheFile(FileName)

            print(FILE) # Display the information of the file

        elif Command.lower() == "save":

            # Input the file that want to be openned
            FileName = str(input("Input the file name to begin saving file (you may include the file extension): "))

            if FileName.lower() == "end":
                break


            FILE = gcv.FileAttribute.OpenTheFile(FileName)

            # Input the file name to save
            SaveFile = str(input("Input the filename that you want to save : "))

            SavingRow = str(input("Input the row that you want to be save"))

            Format = str(input("Input the format that you want to save (.xlsx, .csv)"))

            gcv.FileAttribute.SaveTheFile(SaveFile, "csv", FILE.iloc[[SavingRow]], False)

            # Get the information for saving the file
            print("Finish Saving File")

        elif Command.lower() == "end":
            break

        


        