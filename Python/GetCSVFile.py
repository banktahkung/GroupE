import pandas as pd
import os 
import time
    
class Checking:
    # Add the folder "Inventory" if it is not existed
    def CheckingTheDefaultEnvironment(path:str):
        print("Checking the directory (Inventory)")

        if not (os.path.exists("Inventory")):
            os.makedirs("Inventory")
    
    # Checking the extension of the "existed file":
    def CheckingTheFileExtension(filename:str):

        # Build up the path to the file
        PathFileName = f"Inventory/{filename}"

        # If the extension is not include in the given filename
        if (".csv" not in filename or ".xlsx" not in filename):

            # If the file has csv extension
            if (os.path.isfile(PathFileName + ".csv")):
                return f"{PathFileName}.csv"

            # If the file has xlsx extension
            elif (os.path.isfile(PathFileName + ".xlsx")):
                return f"{PathFileName}.xlsx"

        return PathFileName


class FileAttribute:
    # Open the given file (in excel format)
    def OpenTheCSVFile(file_name:str):

        # The main file 
        global FILE_UNIVERSAL

        # Checking whether the current folder contain the "Inventory folder"
        Checking.CheckingTheDefaultEnvironment("Inventory")

        OriginFileName = Checking.CheckingTheFileExtension(filename=file_name)

        print(f"{OriginFileName}")

        # If the file exist, then import it
        if (os.path.isfile(OriginFileName)):
            print("Found the file")

            print("The content of the file : ")
            
            df = pd.read_csv(OriginFileName)
            print("File opened successfully")

            # First 5 rows
            print(df.head(5))

            FILE_UNIVERSAL = df

        else:
            print("File not found. Did you import it yet??")

        return 0

# Run this if want to test this script
if __name__ == "__main__" :
    print("Running the GET CSV FILE script")
    FileAttribute.OpenTheCSVFile("database101.csv")
    Row_selected_File_Universal = FILE_UNIVERSAL.iloc[[4]]
    print(Row_selected_File_Universal)
    Row_selected_File_Universal['Original_File_Name'] = "database101.csv"
    Row_selected_File_Universal['Row_Index'] = 5
    Row_selected_File_Universal.to_csv(f"Inventory/prototype.csv", index=False)

    FileAttribute.OpenTheCSVFile("prototype.csv")

    last_column = FILE_UNIVERSAL.columns[-5:-1]
    print(f"The last column '{last_column}' value of the selected row is:")
    print(FILE_UNIVERSAL[last_column])
