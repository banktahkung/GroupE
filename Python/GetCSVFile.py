import pandas as pd
import os
import base64
import subprocess
import importlib.util
import sys

HEADER = b'HASHEDFILE'


class Checking:

    # Add the folder "Inventory" if it is not existed
    @staticmethod
    def CheckingTheDefaultEnvironment(path: str):
        """
        Checking whether the current path contain the folder "path" to store the file or not.
        If not, create the "path" folder to store the information

        Parameter(s):
        - path (str) : Path to the folder

        """

        global CURRENT_PATH, DOCUMENT_FOLDER

        if not (os.path.exists(path)):
            os.makedirs(path)

        # Get the current path of this file
        CURRENT_PATH = os.path.dirname(os.path.abspath(__file__))
        DOCUMENT_FOLDER = path

    # Checking the extension of the "existed file":
    @staticmethod
    def CheckingTheFileExtension(filename: str):
        """
        Check if the given filename includes an extension. If not, append the filename with ".csv" or ".xlsx" and check if the file exists.
        If found, return the filename with the path. If not found, return the original filename.

        Parameters:
        - filename (str): The file name that may or may not contain an extension.

        Returns:
        - str: The complete file name with extension if it exists, or the original file name with the path.
        """

        global CURRENT_PATH, DOCUMENT_FOLDER

        # Build up the path to the file
        PathFileName = os.path.join(CURRENT_PATH, DOCUMENT_FOLDER, filename)

        # If the extension is not include in the given filename
        if (".csv" not in filename or ".xlsx" not in filename):

            # If the file has csv extension
            if (os.path.isfile(PathFileName + ".csv")):
                return f"{PathFileName}.csv"

            # If the file has xlsx extension
            elif (os.path.isfile(PathFileName + ".xlsx")):
                return f"{PathFileName}.xlsx"

        return PathFileName

    # Checking whether the python module is already install or not
    def CheckingTheModule(module_str: str = None, module_list: list = []):
        """
            Checking if the given module is already included in the environment or not. If not then install.

            Parameters:
            - module_str (str): The module name that need to be check and install if have not installed yet
            - module_list (list) : The list of module name that need to be check and install if have not installed yet (Default : [] (empty list))
        """

        # Install and import the module if the module isn't already installed yet
        def install_and_import(package):
            spec = importlib.util.find_spec(package)
            if spec is None:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            globals()[package] = importlib.import_module(package)
        
        # If the input moudle_str is not None (No input)
        if module_str is not None:
            install_and_import(module_str)
        
        # Install the module from the list
        for module in module_list:
            install_and_import(module)


class FileAttribute:

    # Open the given file (in excel format)
    @staticmethod
    def OpenTheCSVFile(filename: str):
        """
        Open the given file name (if the given file name doesn't contain the extension then append it with ".csv" or ".xlsx")

        Parameters:
        - filename (str): The name of file

        Return:
        - pd.DataFrame :  return the content of the given file. if the file doesn't existed, return "None"
        """

        # Find the closest file name in the same folder
        def MostFileFounder(filename:str):

            """
            Find the closest file name in the same folder based on Levenshtein distance.

            Parameters:
            - filename (str): Filename to find closest match for.

            Returns:
            - str: Closest filename(s) separated by comma.
            """

            def levenshtein_distance(s1, s2):
                """
                Compute Levenshtein distance between two strings.

                Parameters:
                - s1 (str): First string.
                - s2 (str): Second string.

                Returns:
                - int: Levenshtein distance between s1 and s2.
                """

                if len(s1) > len(s2):
                    s1, s2 = s2, s1

                distances = range(len(s1) + 1)

                for i2, c2 in enumerate(s2):
                    distances_ = [i2 + 1]
                    for i1, c1 in enumerate(s1):
                        if c1 == c2:
                            distances_.append(distances[i1])
                        else:
                            distances_.append(1 + min((distances[i1], distances[i1 + 1], distances_[-1])))
                    distances = distances_

                return distances[-1]

            folder_path = os.path.dirname(filename)
            file_name = os.path.basename(filename)

            print(folder_path)

            closest_filenames = []
            closest_distance = float('inf')

            for file in os.listdir(folder_path):
                if file != file_name:
                    distance = levenshtein_distance(file_name, file)

                    if distance < closest_distance:
                        closest_filenames = [file]
                        closest_distance = distance
                    elif distance == closest_distance:
                        closest_filenames.append(file)

            return ', '.join(closest_filenames)


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
            raise FileNotFoundError(
                f"File {filename} not found. Did you import it yet? \nSearching Path : {OriginFileName} \n Do you mean : {MostFileFounder(filename=OriginFileName)}")

    # Save the Excel file
    @staticmethod
    def SaveTheFile(filename: str = "test" , format: str = "csv", Information: pd.DataFrame = None, hash: bool = False):
        """
        Save the given information to a file in the specified format.

        Parameters:
        - filename (str): The name of the file to save (without extension) [csv, xlsx].
        - format (str): The format to save the file in ('csv' or 'xlsx').
        - information (pd.DataFrame): The data to be saved.
        - hash (bool) : Whether to hash the file or not. Default is 'False'.

        """

        # If the given information is not None
        if Information is not None:
            TempFileName = os.path.join(CURRENT_PATH, DOCUMENT_FOLDER, filename + f"_temp.{format.lower()}")

            # Export the information as the given extension
            if format.lower() == 'csv':
                Information.to_csv(TempFileName, index=False)

            elif format.lower() == 'xlsx':
                Information.to_excel(TempFileName, index=False)

            # If hash is True, then perform hash to the information
            if hash:
                with open(TempFileName, "rb") as file:
                    data = file.read()
                encoded_data = HEADER + base64.b64encode(data)
                
                # Create the file and rewrite it with the encoded data
                with open(os.path.join(CURRENT_PATH, DOCUMENT_FOLDER, filename + f".{format.lower()}"), "wb") as file:
                    file.write(encoded_data)

                os.remove(TempFileName)
            else:
                os.rename(TempFileName, os.path.join(CURRENT_PATH, DOCUMENT_FOLDER, filename + f".{format.lower()}"))
        else:
            raise ValueError("The given information cannot be empty")


# Run this script if want to test this script
if __name__ == "__main__":
    print("Running the GET CSV FILE script")

    Checking.CheckingTheModule("pandas")

    # Modules to check/install
    modules_to_install = ["pandas", "openpyxl"]
    Checking.CheckingTheModule(module_list=modules_to_install)

    # Replace the parameter with your "filename" (you may include the extension)
    FILE_UNIVERSAL = FileAttribute.OpenTheCSVFile("data")

    # Select the forth row of the csv file
    Row_selected_File_Universal = FILE_UNIVERSAL.iloc[[4]]
    print(Row_selected_File_Universal)

    # Add the column "Original_File_Name" and "Row_Index"
    Row_selected_File_Universal['Original_File_Name'] = "database101.csv"
    Row_selected_File_Universal['Row_Index'] = 5

    # Save the file
    FileAttribute.SaveTheFile("prototype", "xlsx", Row_selected_File_Universal)

    # Open the file
    FILE_UNIVERSAL = FileAttribute.OpenTheCSVFile("prototype.csv")

    # Get the last 5 columns of the file
    last_column = FILE_UNIVERSAL.columns[-5:-1]

    print(f"The last column '{last_column}' value of the selected row is:")
    print(FILE_UNIVERSAL[last_column])
