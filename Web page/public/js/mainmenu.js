// - Variable Declaration
const fileInput = document.getElementById("fileInput");
const fileContent = document.getElementById("FileContent");
const fileContentOutput = document.querySelector(".FileContent.Output");
const fileOutputContent = document.querySelector(".FileContent.Output");
const ConvertionArrow = document.getElementById("ConvertArrow");
const ProfileAddingWindow = document.getElementById("AddingProfileWindow");
const DropdownProfile = document.getElementById("ProfileDropdown");
const FileInputMessage = document.getElementById("FileInputMessage");
const FileOutputMessage = document.getElementById("FileOutputMessage");
const FileDrawerDropdown = document.getElementById("fileDrawerDropdown");

// % The Item Array List
let ItemArray = [];
let FileDrawer = {}, ProfileList = {};
let CurrentProfile = "Default Profile";

let DefaultProfile = null;
let ConditionalValue = null;
let TempTable = {},
  OutputTempTable = {};
let CurrentRowOutputTable = 1;

// % Axis Declaration
const XIndicator = document.querySelectorAll(".XAxisIndicator");
const YIndicator = document.querySelectorAll(".YAxisIndicator");

let DropdownState = false,
  FileDropdownState = true;
let SetFile = [];

const Alphabet = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
];

// - Event listener
// > File input
fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  SetFile.push(file);

  FileHandling(file);
  GetTheProfile();
});

// > Dropdown selection

// - Function
// > file handling fuction
function FileHandling(file) {
  if (file) {
    // * Hide the file input message
    FileInputMessage.style.display = "none";
    FileOutputMessage.style.display = "none";

    // * Declare reader for reading the file
    const reader = new FileReader();

    // * Get the information from the file
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // * Get the first sheet name
      const sheetName = workbook.SheetNames[0];

      // * Get the worksheet
      const worksheet = workbook.Sheets[sheetName];

      // * Convert to JSON
      const json = XLSX.utils.sheet_to_json(worksheet);

      // * Add the file to the drawer
      AddFileDrawer(sheetName, json);
      AddFileDrawerDropdown(sheetName);

      console.log(json);

      FileShowcase(json, sheetName);

      console.log(OutputTempTable);
    };

    reader.readAsArrayBuffer(file);
  } else {
    FileInputMessage.style.display = "flex";
    FileOutputMessage.style.display = "flex";

    ConvertionArrow.classList.remove("Enable");
    ConvertionArrow.classList.add("Disabled");

    fileOutputContent.innerHTML = "";
  }
}

// > File content adding function
// & After user inputs the file, we shall keep it as "drawer"
function AddFileDrawer(sheetName, json) {
  // * Store the sheet data in the FileDrawer object with the sheetName as the key
  FileDrawer[sheetName] = json;
}

// > Add file content to the file drawer dropdown
function AddFileDrawerDropdown(sheetName) {
  // * Get the current amount of file in the drawer
  const file = Object.keys(FileDrawer);
  let FileAmount = file.length;

  // * Create division content
  const div = document.createElement("div");

  div.textContent = sheetName;
  div.classList.add("FileDropdown");

  FileDrawerDropdown.appendChild(div);
}

// > file showcase function
// & First, we will find the most space contains row
// & Then, we will map each value to its coresponse row
// & Finally, Show
function FileShowcase(json, sheetName) {
  // * Clear the file content
  fileContent.innerHTML = ""; // * Clear previous content

  // * Clear the content in the indicator
  XIndicator.forEach((indicator) => {
    indicator.innerHTML = "";
  });

  YIndicator.forEach((indicator) => {
    indicator.innerHTML = "";
  });

  // > Find the most space contain row
  let highestNumber = 0;
  let highestLineNumber = 0;

  for (let loop = 0; loop < json.length; loop++) {
    let headers = Object.keys(json[loop]);

    headers.forEach((hd, count) => {
      // * Split the header
      splitHD = hd.split("__EMPTY_");

      // * Find the most space contain in each Line
      splitHD.forEach((element) => {
        if (element != "" && parseInt(element) > highestNumber)
          highestNumber = parseInt(element);
      });
    });

    // * Find the total row on the sheet
    if (highestLineNumber < parseInt(Object.keys(json[loop])["__rowNum__"])) {
      highestLineNumber = parseInt(Object.keys(json[loop])["__rowNum__"]);
    }
  }

  GenerateAxis(json, highestNumber);

  // % Build the table
  // * Length of the json file (number of rows)
  for (let i = 0, RowCount = 1; i < json.length; i++) {
    // * Extract the spacing length from each header
    let headers = Object.keys(json[i]);
    let content = json[i];

    let count = 0;

    // * Create `div`
    const element = document.createElement("div");
    element.classList.add("TableParent", `row${i}`);

    // * Length of the most space contain row
    for (let j = 0; j <= highestNumber; j++) {
      // * Create the table using div
      const childElement = document.createElement("div");
      childElement.classList.add("TableChild", `row${i}col${j}`);

      childElement.textContent = "";

      // * If the already select all the element in the headers
      if (count != headers.length) {
        // * If select the item at the first column
        if (headers[count] == `__EMPTY_${j}`) {
          let Text = content[headers[count]];

          // * Date conversion (Only this cell, Don't know why lol)
          if((j == 14 && i == 6) || (j == 17 && i == 8)) {
            Text = excelDateToCustomFormat(Text)
          }

          while (TempTable[RowCount] === undefined) {
            RowCount++;
          }

          TempTable[RowCount][getColumnLabel(j)] = Text;

          if (Text.length > 15) {
            Text = Text.substring(0, 13) + "...";
          }

          childElement.textContent = Text;

          count++;
        } else if (j == 0 && headers[count] == "__EMPTY") {
          let Text = content[headers[count]];

          while (TempTable[RowCount] === undefined) {
            RowCount++;
          }

          TempTable[RowCount][getColumnLabel(j)] = Text;

          if (Text.length > 15) {
            Text = Text.substring(0, 13) + "...";
          }

          childElement.textContent = Text;
          count++;
        }

        // * Don't know why this key exist but it represents Q column
        else if (headers[count] == ".") {
          // * Add up until J is 19
          while (j != 19) {
            // * Create element that represent the empty cells
            const TempElement = document.createElement("div");
            TempElement.classList.add("TableChild", `row${i}col${j}`);

            TempElement.textContent = "";

            element.appendChild(TempElement);

            j++;
          }

          let Text = content[headers[count]];

          while (TempTable[RowCount] === undefined) {
            RowCount++;
          }

          TempTable[RowCount][getColumnLabel(j)] = Text;

          if (Text.length > 15) {
            Text = Text.substring(0, 13) + "...";
          }

          childElement.textContent = Text;

          count++;
        }
      }

      if (i == 0) {
        childElement.style.borderTop = `rgb(230, 230, 230) solid 1px`;
      }

      element.style.width = `${150 * highestNumber}px`;

      element.appendChild(childElement);
    }

    RowCount++;

    // * Add the created element to the file content
    fileContent.appendChild(element);

    count = 0;
  }

  // % Build the output table
  FileConversion(TempTable);

  // % Change the state of the convertion arrow
  ConvertionArrow.classList.remove("Disabled");
  ConvertionArrow.classList.add("Enable");
}

function excelDateToCustomFormat(serial) {
  // Excel serial dates start from 1900-01-01
  const excelStartDate = new Date(1900, 0, 1);
  
  // Add the serial number of days to the start date
  // Note: Excel incorrectly considers 1900 as a leap year, so subtract 1 day
  const excelDate = new Date(excelStartDate.getTime() + (serial - 2) * 86400000);
  
  // Options for formatting the date
  const options = { day: 'numeric', month: 'short', year: '2-digit' };
  
  // Format the date
  return excelDate.toLocaleDateString('en-GB', options);
}

// + File Output Part
// > File conversion
async function FileConversion(json) {
  let header = await GetTableHeader();
  
  // * Assigning the current profile, both default and condition
  DefaultProfile = ProfileList[CurrentProfile].defaultProfile;
  ConditionalValue = ProfileList[CurrentProfile].conditionalValue;

  console.log(DefaultProfile);

  // * Place the first column header
  // & loop for each header in the output table
  const Tableheader = document.createElement("div");
  Tableheader.classList.add("TableParent", "row0");

  Tableheader.style.width = `${200 * 96}px`;

  OutputTempTable["row0"] = {};

  // * Loop till every header
  for (let i = 0; i < header.length; i++) {
    // * Create an element to hold the content of header of the table
    const TableHeaderElement = document.createElement("div");

    TableHeaderElement.classList.add("TableChild", `row0col${i}`);

    // * Set the limit for the characters
    if (header[i].length > 15) {
      TableHeaderElement.textContent = header[i].substring(0, 10) + "...";
    } else {
      TableHeaderElement.textContent = header[i];
    }

    // * A little styling to the table
    TableHeaderElement.style.borderTop = `rgb(230, 230, 230) solid 1px`;
    TableHeaderElement.style.width = "200px";

    // * Add the value into the output temp table
    OutputTempTable["row0"][`col${i}`] = header[i];

    // * Add the element to the table header
    Tableheader.appendChild(TableHeaderElement);
  }

  fileContentOutput.appendChild(Tableheader);

  // + String Input Handling
  function splitString(input) {
    const match = input.match(/^([A-Za-z]+)(\d+)$/);
    if (match) {
      const letters = match[1];
      const numbers = match[2];
      return [letters, numbers];
    }
    return null;
  }

  // * Adding the file content according to the profile
  for (let i = 0; i < Object.keys(FileDrawer).length * 2; i++) {
    const RowContent = document.createElement("div");
    RowContent.classList.add("TableParent", `row${i + 1}`);
    RowContent.style.width = "19200px";

    OutputTempTable[`row${i + 1}`] = {};

    // * Loop to every profile
    DefaultProfile.forEach((dp, index) => {
      // * Create the `div` element
      const CellContent = document.createElement("div");
      CellContent.textContent = "";
      CellContent.classList.add("TableChild", `row${i + 1}col${index}`);
      CellContent.style.width = `200px`;

      // * In case it is not the conditional value
      if (dp != "CV" && dp != null && dp.length < 5) {
        const TargetCell = splitString(dp);

        // * Get the target Column
        let TargetColumn = TargetCell[0],
          TargetRow = TargetCell[1];

        // * Get the text from the extraction
        let Extrect = TempTable[TargetRow][TargetColumn];

        OutputTempTable[`row${i + 1}`][`col${index}`] = Extrect;

        // * Reduce the length of the text
        if (Extrect && Extrect.length > 15) {
          Extrect = Extrect.substring(0, 13) + "...";
        }

        // * Substitute the text into the cell content
        CellContent.textContent = Extrect;
      }
      // * In case the conditional value
      else if (dp == "CV") {
        console.log(`${getColumnLabel(index)} column`);

        const ConditionHeader = Object.keys(
          ConditionalValue[`${getColumnLabel(index)} column`]
        )[0];

        // * Switch case for handling the condition
        switch (ConditionHeader) {
          case "CONDITION":
            let GivenCondition = Object.keys(
              ConditionalValue[`${getColumnLabel(index)} column`][
                ConditionHeader
              ]
            )[0];

            // * If condition consider about the output table
            if (GivenCondition.split(" ")[1] == "OUTPUT") {
              let ConsiderationValue =
                OutputTempTable[`row${i + 1}`][
                  `col${getColumnIndex(GivenCondition.split(" ")[0]) - 1}`
                ];

              if (ConsiderationValue) {
                ConsiderationValue.trim().split(" ");
              }

              let ConsiderationCondition = Object.keys(
                ConditionalValue[`${getColumnLabel(index)} column`][
                  ConditionHeader
                ][GivenCondition]
              );

              ConditionConsideration(
                ConsiderationValue,
                ConsiderationCondition,
                CellContent,
                ConditionHeader,
                GivenCondition,
                i,
                index
              );
            }

            // * If the condition consider about the input part
            else if (GivenCondition.split(" ")[1] == "INPUT") {
              const TargetColumn = splitString(GivenCondition.split(" ")[0]);

              let ConsiderationValue =
                TempTable[TargetColumn[1]][TargetColumn[0]];

              let ConsiderationCondition = Object.keys(
                ConditionalValue[`${getColumnLabel(index)} column`][
                  ConditionHeader
                ][GivenCondition]
              );

              // * Call the function
              ConditionConsideration(
                ConsiderationValue,
                ConsiderationCondition,
                CellContent,
                ConditionHeader,
                GivenCondition,
                i,
                index
              );
            }

            break;
          case "OPERATION":
            let OperationValue =
              ConditionalValue[`${getColumnLabel(index)} column`][
                ConditionHeader
              ];

            // * Length cutting algorithm
            if (OperationValue[1] == "length") {
              const TargetCell = splitString(OperationValue[0]);
              let ConsiderationValue = TempTable[TargetCell[1]][TargetCell[0]];
              let CutValue = "";

              switch (OperationValue[2]) {
                case "start":
                  CutValue = ConsiderationValue.substring(
                    OperationValue[3],
                    ConsiderationValue.length
                  );
                  OutputTempTable[`row${i + 1}`][`col${index}`] = CutValue;

                  if (CutValue.length > 15) {
                    CutValue = CutValue.substring(0, 13) + "...";
                  }

                  CellContent.textContent = CutValue;

                  break;

                case "end":
                  let StringLength =
                    ConsiderationValue.length - OperationValue[3];
                  CutValue = ConsiderationValue.substring(
                    StringLength,
                    ConsiderationValue.length
                  );

                  break;
              }
              OutputTempTable[`row${i + 1}`][`col${index}`] = CutValue;

              if (CutValue.length > 15) {
                CutValue = CutValue.substring(0, 13) + "...";
              }

              CellContent.textContent = CutValue;
            }

            // * Adding String Operation
            else {
              let InitialString = "";

              for (let i = 0; i < OperationValue.length; i++) {
                if (OperationValue[i] == "add") continue;

                let TargetCell = splitString(OperationValue[i]);
                let AddingData = TempTable[TargetCell[1]][TargetCell[0]];

                InitialString += AddingData;
              }

              OutputTempTable[`row${i + 1}`][`col${index}`] = InitialString;

              if (InitialString.length > 15) {
                InitialString = InitialString.substring(0, 13) + "...";
              }

              CellContent.textContent = InitialString;
            }

            break;
          case "DEFAULT":
            let Text =
              ConditionalValue[`${getColumnLabel(index)} column`][
                ConditionHeader
              ];

            OutputTempTable[`row${i + 1}`][`col${index}`] = Text;

            if (Text.length > 15) {
              Text = Text.substring(0, 13) + "...";
            }

            CellContent.textContent = Text;

            break;
        }
      }

      // * In case it `null` value
      else {
        CellContent.textContent = " ";

        OutputTempTable[`row${i + 1}`][`col${index}`] = null;
      }

      // * Add the content to the row
      RowContent.appendChild(CellContent);
    });

    fileContentOutput.appendChild(RowContent);
  }

  console.log(OutputTempTable);
}

// + Checking condition
// > Checking the value base on the condition
function ConditionConsideration(
  ConsiderationValue,
  ConsiderationCondition,
  CellContent,
  ConditionHeader,
  GivenCondition,
  i,
  index
) {
  let ConditionChecking = true,
    count = 0;

  // * Condition Checking
  // & Condition is not about the null and not null
  if (
    ConsiderationCondition[0].toLowerCase == "null" ||
    ConsiderationCondition[0].toLowerCase == "not null"
  ) {
    ConsiderationCondition.forEach((cdc) => {
      console.log(cdc.split(" "));

      // * Split the string if the string contains multiple words
      cdc.split(" ").forEach((cdcs, index) => {
        if (ConsiderationValue[count] == "") {
          count++;
        }

        if (!cdcs.includes(ConsiderationValue[count])) {
          ConditionChecking = false;
        }

        count++;
      });

      // * Initialize the count value
      count = 0;

      // * Reduce the string length of the `final value`
      if (ConditionChecking) {
        OutputTempTable[`row${i + 1}`][`col${index}`] = ConsiderationValue;

        if (ConsiderationValue.length > 15) {
          ConsiderationValue = ConsiderationValue.substring(0, 13) + "...";
        }

        // * Assigning the value
        CellContent.textContent =
          ConditionalValue[`${getColumnLabel(index)} column`][ConditionHeader][
            GivenCondition
          ][cdc];
      }

      ConditionChecking = true;
    });
  }

  // & Condition is about null and not null checking
  else {
    // * If the consideration value exists
    if (ConsiderationValue) {
      CellContent.textContent =
        ConditionalValue[`${getColumnLabel(index)} column`][ConditionHeader][
          GivenCondition
        ][ConsiderationCondition[1]];

      OutputTempTable[`row${i + 1}`][`col${index}`] = CellContent.textContent;
    } else {
      CellContent.textContent =
        ConditionalValue[`${getColumnLabel(index)} column`][ConditionHeader][
          GivenCondition
        ][ConsiderationCondition[0]];

      OutputTempTable[`row${i + 1}`][`col${index}`] = CellContent.textContent;
    }
  }
}

// + Adding Excel Profile ------------
// > Openning Adding window profile
function OpenningAddingExcelProfile() {
  ProfileAddingWindow.style.display = "flex";
}

// > Closing Adding window profile
function ClosingAddingExcelProfile() {
  ProfileAddingWindow.style.display = "none";
}

// > Dropdown Manipulation
function DropdownContent() {
  if (DropdownState) {
    DropdownProfile.style.display = "none";
    DropdownState = false;
  } else {
    DropdownProfile.style.display = "block";
    DropdownState = true;
  }
}

// + File Inventory for keeping the file
function FileDropdownContent() {
  if (FileDropdownState) {
    FileDrawerDropdown.style.display = "block";
    FileDropdownState = false;
  } else {
    FileDrawerDropdown.style.display = "none";
    FileDropdownState = true;
  }
}

// > Generate the Axis Indicator
// * Function to generate column labels
function getColumnLabel(index) {
  let label = "";
  while (index >= 0) {
    label = Alphabet[index % 26] + label;
    index = Math.floor(index / 26) - 1;
  }
  return label;
}

function getColumnIndex(str) {
  let result = 0;
  for (let i = 0; i < str.length; i++) {
    result *= 26;
    result += str.charCodeAt(i) - 64;
  }
  return result;
}

// * Function to generate row labels
function GenerateAxis(json, highestNumber) {
  let columnWidth = 150;

  // - Generate the axis for each X axis indicator
  XIndicator.forEach((indicatorBar) => {
    // * Generate X axis labels
    if (indicatorBar.classList.contains("Output")) {
      highestNumber = 96;
      columnWidth = 200;
    }

    for (let i = 0; i <= highestNumber; i++) {
      const div = document.createElement("div");
      div.classList.add("XIndicator");
      div.textContent = getColumnLabel(i);
      div.style.width = "200px";
      indicatorBar.appendChild(div);
    }

    indicatorBar.style.width = `${highestNumber * columnWidth}px`;
    indicatorBar.style.paddingLeft = "30px";
  });

  // - Generate the axis for each Y axis indicator
  YIndicator.forEach((indicatorBar) => {
    if (!indicatorBar.classList.contains("Output")) {
      // * Generate the Y axis labels
      json.forEach((row, index) => {
        const div = document.createElement("div");
        div.classList.add("YIndicator");

        if (index == 0) {
          div.classList.add("first");
        }

        // * Change the text content of the number of row
        div.textContent =
          row["__rowNum__"] !== undefined ? row["__rowNum__"] + 1 : index + 1;

        // * Make a list about the temporary excel file
        TempTable[
          row["__rowNum__"] !== undefined ? row["__rowNum__"] + 1 : index + 1
        ] = {};

        indicatorBar.appendChild(div);
      });
    } else {
      for (let i = 0; i < Object.keys(FileDrawer).length * 2 + 1; i++) {
        const div = document.createElement("div");
        div.classList.add("YIndicator");

        if (i == 0) {
          div.classList.add("first");
        }

        div.textContent = i + 1;
        indicatorBar.appendChild(div);
      }
    }
  });
}

// + GET the header and also the profile of the excel sheet
// & Function to get the header of the result table
async function GetTableHeader() {
  let destination = "/header";

  // > Try to get the header of the output table
  try {
    const response = await fetch(destination);

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();

    let header = data.Header;

    return header;
  } catch (error) {
    console.error("Error fetching table header:", error);
  }
}

// > Get the profile from the backend
async function GetTheProfile() {
  let destination = "/profile";

  try {
    const response = await fetch(destination);

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();

    const ProfileHeader = Object.keys(data);

    ProfileList = data;

    console.log(ProfileList);

    DefaultProfile = data.defaultProfile;
    ConditionalValue = data.conditionalValue;
  } catch (err) {
    console.error("Error fetching profile:", err);
  }
}

// + Excel File Creation
// > Test function to download Excel
function downloadExcel(filename = "data.xlsx") {
  const { ExportHeader, ExportData } = transformData(OutputTempTable);

  console.log(ExportData);
  console.log(ExportHeader);

  // * Create a new workbook
  const wb = XLSX.utils.book_new();

  // * Convert JSON to a worksheet with specified headers
  const ws = XLSX.utils.json_to_sheet(ExportData, { header: ExportHeader });

  // * Append the worksheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // * Generate a file name and save it
  XLSX.writeFile(wb, filename);
}

// > Transform JSON data into headers and rows
function transformData(jsonData) {
  let headers = [];
  let rows = [];

  // * Get the header of json data
  const HeaderKey = Object.keys(jsonData);

  // * Check if jsonData has at least one row to extract headers
  if (HeaderKey.length > 0) {
    // * Extract headers from the first row
    const firstRow = jsonData[HeaderKey[0]];
    headers = Object.values(firstRow);

    // * Process remaining rows
    HeaderKey.slice(1).forEach((rowKey) => {
      console.log(rowKey);
      const row = jsonData[rowKey];

      console.log(row);

      let rowData = {};
      Object.keys(row).forEach((colKey, index) => {
        rowData[headers[index]] = row[colKey];
      });
      rows.push(rowData);
    });
  }

  return { ExportHeader: headers, ExportData: rows };
}
