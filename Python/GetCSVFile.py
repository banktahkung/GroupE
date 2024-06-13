import pandas as pd
import os 
import base64

HEADER = b'HASHEDFILE'
    
class Checking:

    # Add the folder "Inventory" if it is not existed
    @staticmethod
    def CheckingTheDefaultEnvironment(path:str):
        """
        Checking whether the current path contain the folder "path" to store the file or not.
        If not, create the "path" folder to store the information

        Parameter(s):
        - path (str) : Path to the folder

        """

        if not (os.path.exists(path)):
            os.makedirs(path)
    
    # Checking the extension of the "existed file":
    @staticmethod
    def CheckingTheFileExtension(filename:str):
        """
        Check if the given filename includes an extension. If not, append the filename with ".csv" or ".xlsx" and check if the file exists.
        If found, return the filename with the path. If not found, return the original filename.

        Parameters:
        - filename (str): The file name that may or may not contain an extension.

        Returns:
        - str: The complete file name with extension if it exists, or the original file name with the path.
        """

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
    @staticmethod
    def OpenTheCSVFile(filename:str):
        """
        Open the given file name (if the given file name doesn't contain the extension then append it with ".csv" or ".xlsx")

        Parameters:
        - filename (str): The name of file

        Return:
        - pd.DataFrame :  return the content of the given file. if the file doesn't existed, return "None"
        """

        # Checking whether the current folder contain the "Inventory folder"
        Checking.CheckingTheDefaultEnvironment("Inventory")

        OriginFileName = Checking.CheckingTheFileExtension(filename=filename)

        # If the file exist, then import it
        if (os.path.isfile(OriginFileName)):
            
            # Open the file
            with open(OriginFileName, "rb") as file:
                data = file.read()

            # Check whether the file is hashed or not
            # If hash, then decode it
            if data.startswith(HEADER):
                DecodedData = base64.b64decode(data[len(HEADER):])

                # Write the data as the decoded data to the temporary csv file
                with open(OriginFileName + "_temp.csv", "wb") as temp_file:
                    temp_file.write(DecodedData)

                # Read the data
                df = pd.read_csv(OriginFileName + "_temp.csv")

                # Remove the temporary file
                os.remove(OriginFileName + "_temp.csv")

            # If not hash, then read it as csv
            else:
                df = pd.read_csv(OriginFileName)

            return df

        else:
            raise FileNotFoundError(f"File {filename} not found. Did you import it yet? \nSearching Path : {OriginFileName}")

    # Save the Excel file
    @staticmethod
    def SaveTheFile(filename:str = "test", format:str = "csv", Information:pd.DataFrame = None, hash:bool = False):
        """
        Save the given information to a file in the specified format.

        Parameters:
        - filename (str): The name of the file to save (without extension) [csv, xlsx].
        - format (str): The format to save the file in ('csv' or 'xlsx').
        - information (pd.DataFrame): The data to be saved.
        - hash (bool) : Whether to hash the file or not. Default is 'False'.

        """

        if Information:

            # Save the file in "csv" format
            if ("csv" in format.lower()):

                # Save the information to the temporary "csv" file
                TempFileName = f"Inventory/{filename}_temp.csv"
                Information.to_csv(TempFileName, index=False)

                if hash:

                    # Open the temporary file and retrieve the data to perform hash
                    with open(TempFileName, "rb") as file:
                        data = file.read()

                    # Encoded the data
                    encoded_data = HEADER + base64.b64encode(data)

                    # Write the encoded data to the "filename".csv
                    with open(f"Inventory/{filename}.csv", "wb") as file:
                        file.write(encoded_data)

                    # Remove the temporary file
                    os.remove(TempFileName)
                else:

                    # Rename the file if the hash value is "False"
                    os.rename(TempFileName, f"Inventory/{filename}.csv")

            
            # Save the file in "xlsx" format
            elif ("xlsx" in format.lower()):

                # Save the information to the temporary "xlsx" file
                TempFileName = f"Inventory/{filename}_temp.xlsx"
                Information.to_excel(TempFileName, index=False)

                if hash:

                    # Open the temporary file and retrieve the data to perform hash
                    with open(TempFileName, "rb") as file:
                        data = file.read()

                    # Encoded the data
                    encoded_data = HEADER + base64.b64encode(data)

                    # Write the encoded data to the "filename".csv
                    with open(f"Inventory/{filename}.xlsx", "wb") as file:
                        file.write(encoded_data)

                    # Remove the temporary file
                    os.remove(TempFileName)
                else:

                    # Rename the file if the hash value is "False"
                    os.rename(TempFileName, f"Inventory/{filename}.xlsx")
        
        else:
            raise ValueError("The given information can't be empty")

# Run this script if want to test this script
if __name__ == "__main__" :
    print("Running the GET CSV FILE script")

    # Replace the "database101.csv" with your filename (you may include the extension)
    FILE_UNIVERSAL = FileAttribute.OpenTheCSVFile("a.csv")

    # Select the forth row of the csv file
    Row_selected_File_Universal = FILE_UNIVERSAL.iloc[[4]]
    print(Row_selected_File_Universal)

    # Add the column "Original_File_Name" and "Row_Index"
    Row_selected_File_Universal['Original_File_Name'] = "database101.csv"
    Row_selected_File_Universal['Row_Index'] = 5

    # Save the file
    FileAttribute.SaveTheFile(filename="prototype", format="csv", hash=True)

    # Open the file
    FILE_UNIVERSAL = FileAttribute.OpenTheCSVFile("prototype.csv")

    # Get the last 5 columns of the file
    last_column = FILE_UNIVERSAL.columns[-5:-1]

    print(f"The last column '{last_column}' value of the selected row is:")
    print(FILE_UNIVERSAL[last_column])
