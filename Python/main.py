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

        # Get the input file name
        FileName = str(input("Input the file name the display (you may include the file extension)): "))

        if (FileName == "end"):
            break

        # Openning the file
        gcv.FileAttribute.OpenTheCSVFile(FileName)
        