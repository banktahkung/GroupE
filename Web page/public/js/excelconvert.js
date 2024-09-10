// % TMR TASK
// % MULTIPLE EXCEL FILE MANIPULATION -> be able to select what file should be displayed and then highlight the one that is belonged to the source file

// - Variable Declaration
const fileInput = document.getElementById("fileInput");
const fileContent = document.getElementById("FileContent");
const fileOutputContent = document.querySelector(".FileContent.Output");
const DropdownProfile = document.getElementById("ProfileDropdown");
const FileInputMessage = document.getElementById("FileInputMessage");
const FileOutputMessage = document.getElementById("FileOutputMessage");
const FileDrawerDropdown = document.getElementById("fileDrawerDropdown");
const AddProfileText = document.getElementById("AddProfile");
const ProfileIndicator = document.getElementById("CurrentProfileDeclaration");
const WarningWindow = document.getElementById("WarningWindow");
const WarningMessage = document.getElementById("WarningMessage");

// % The Item Array List
let ItemArray = [];
let FileDrawer = {},
  ProfileList = {},
  TempTable = {},
  OutputTempTable = {};
let CurrentProfile = null,
  DefaultProfile = null,
  ConditionalValue = null,
  FileData = null,
  CurrentFile = null;

let WarningCheck = false,
  WindowExport = false;
let CurrentRowOutputTable = 1;

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

document.addEventListener("DOMContentLoaded", function () {
  const dropArea = document.getElementById("ShowcaseFile");

  // * Add the `click` listener
  dropArea.addEventListener("click", () => {
    if (document.getElementById("FileContent").innerHTML == "") {
      fileInput.click();
    }
  });

  // * Prevent default when drag over the container
  dropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropArea.classList.add("dragover");
  });

  dropArea.addEventListener("dragleave", (e) => {
    e.preventDefault();
    dropArea.classList.remove("dragover");
  });

  // * Handle file drop
  dropArea.addEventListener("drop", (e) => {
    e.preventDefault();

    const files = e.dataTransfer.files;
    dropArea.classList.remove("dragover");

    if (files.length > 0) {
      console.log(files);

      for (let i = 0; i < files.length; i++) {
        SetFile.push(files[i]);

        FileHandling(files[i]);
      }

      GetTheProfile();
    }
  });
});

// - Function
// > file handling fuction
function FileHandling(file) {
  const ConvertionArrow = document.getElementById("ConvertArrow");

  if (file) {
    // * Hide the file input message
    FileInputMessage.style.display = "none";

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

      // * Initialize the Output table
      OutputTempTable = {};

      // * Initialize the Output section
      fileOutputContent.innerHTML = "";

      // * Add the file to the drawer
      AddFileDrawer(file.name, json);
      AddFileDrawerDropdown(file.name);

      // * Show case the file
      FileShowcase(json, file.name);
      FileConversion(TempTable);

      // * Assign the data to global variable (not secure)
      FileData = json;

      console.log(CurrentProfile);
    };

    reader.readAsArrayBuffer(file);
  } else {
    // * Return to the initial state
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
  FileDrawer[sheetName] = {
    Json: json,
    BackgroundColor: RandomBackgroundColor(),
  };
}

// > Add file content to the file drawer dropdown
// & If the file is not existed in the drawer yet, then append it. Otherwise, ignore for prevent duplication of files
function AddFileDrawerDropdown(sheetName) {
  // * Get the current amount of file in the drawer
  const file = Object.keys(FileDrawer);

  let fileContain = false;

  FileDrawerDropdown.childNodes.forEach((cn) => {
    if (cn.textContent == sheetName) fileContain = true;
  });

  if (!fileContain) {
    // * Create division content
    const div = document.createElement("div");

    div.textContent = sheetName;
    div.classList.add("FileDropdown");

    // * Remove the no file division
    if (file.length == 1) {
      FileDrawerDropdown.innerHTML = "";

      div.classList.add("first");
      div.style.borderRadius = `20px`;
      div.style.borderBottom = `none`;
    }
    // * If the length of file drawer is more than 1
    else if (file.length > 1) {
      const firstFileInDrawer = document.querySelector(".FileDropdown.first");
      const lastFileInDrawer = document.querySelector(".FileDropdown.last");

      firstFileInDrawer.style.borderRadius = `20px 20px 0px 0px`;
      firstFileInDrawer.style.borderBottom = `0.5px solid #000`;

      // * Check w
      if (lastFileInDrawer) {
        lastFileInDrawer.classList.remove("last");
      }

      div.classList.add("last");
    }
    div.addEventListener("click", () => {
      DisplayTheFile(div.textContent, FileDrawer[div.textContent.trim()].Json);
    });

    FileDrawerDropdown.appendChild(div);
  }
}

// >  Random background color
function RandomBackgroundColor() {
  const r = Math.floor(Math.random() * 256);
  const g = Math.floor(Math.random() * 256);
  const b = Math.floor(Math.random() * 256);

  return `rgb(${r}, ${g}, ${b}, 0.13)`;
}

// > Select the file to display
function DisplayTheFile(filename, json) {
  FileInputMessage.style.display = "none";

  // * Detect whether the user already change the file or not
  if (filename != CurrentFile) {
    FileShowcase(json, filename);
  }
}

// > file showcase function
// & First, we will find the most space contains row
// & Then, we will map each value to its coresponse row
// & Finally, Show
function FileShowcase(json, sheetName) {
  // * Clear the file content
  fileContent.innerHTML = ""; // * Clear previous content

  // * Find the most space contain row
  let highestNumber = 0;
  let highestLineNumber = 0;

  for (let loop = 0; loop < json.length; loop++) {
    let headers = Object.keys(json[loop]);

    headers.forEach((hd, count) => {
      // * Split the header
      splitHD = hd.split("__EMPTY_");

      // * Find the most space contain in each Line
      splitHD.forEach((element) => {
        if (element != "" && parseInt(element) + 1 > highestNumber)
          highestNumber = parseInt(element) + 1;
      });
    });

    // * Find the total row on the sheet
    if (highestLineNumber < parseInt(Object.keys(json[loop])["__rowNum__"])) {
      highestLineNumber = parseInt(Object.keys(json[loop])["__rowNum__"]);
    }
  }

  CurrentFile = sheetName;

  GenerateInputAxis(json, highestNumber);

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

      childElement.style.backgroundColor =
        FileDrawer[CurrentFile]["BackgroundColor"];

      // * Add hovering effect to the child element
      childElement.addEventListener("mouseover" || "mouseenter", () => {
        childElement.style.backgroundColor =
          FileDrawer[CurrentFile]["BackgroundColor"].substring(
            0,
            FileDrawer[CurrentFile]["BackgroundColor"].length - 4
          ) + "0.35)";
      });

      childElement.addEventListener("mouseleave", () => {
        childElement.style.backgroundColor =
          FileDrawer[CurrentFile]["BackgroundColor"];
      });

      childElement.textContent = "";

      // * If the already select all the element in the headers
      if (count != headers.length) {
        // * If select the item at the first column
        if (
          headers[count].includes(`__EMPTY_${j}`) ||
          (j == 0 && headers[count] == "__EMPTY")
        ) {
          let Text = content[headers[count]];

          // * Date conversion (Only these cells, Don't know why lol)
          if ((j == 14 && i == 6) || (j == 17 && i == 8)) {
            Text = excelDateToCustomFormat(Text);
          }

          while (TempTable[RowCount] === undefined) {
            RowCount++;
          }

          TempTable[RowCount][getColumnLabel(j)] = Text;

          // * Cut the text length
          Text = stringReduction(Text, 15);

          childElement.textContent = Text;

          count++;
        }

        // * Don't know why this key exist but it represents U column
        else if (headers[count] == ".") {
          let skipVaule = j;

          // * Add up until J is 19
          while (j != 19) {
            // * Create element that represent the empty cells
            const TempElement = document.createElement("div");
            TempElement.classList.add("TableChild", `row${i}col${j}`);

            TempElement.textContent = "";
            TempElement.style.backgroundColor =
              FileDrawer[CurrentFile]["BackgroundColor"];

            TempElement.addEventListener("mouseover", () => {
              TempElement.style.backgroundColor =
                FileDrawer[CurrentFile]["BackgroundColor"].substring(
                  0,
                  FileDrawer[CurrentFile]["BackgroundColor"].length - 3
                ) + "0.3)";
            });

            element.appendChild(TempElement);

            j++;
          }

          let Text = content[headers[count]];

          while (TempTable[RowCount] === undefined) {
            RowCount++;
          }

          TempTable[RowCount][getColumnLabel(j)] = Text;

          Text = stringReduction(Text, 15);

          childElement.classList.remove(`row${i}col${skipVaule}`);
          childElement.classList.add(`row${i}col${j}`);
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

  if (!FileDrawer[sheetName]["Convert"]) {
    FileDrawer[sheetName]["Convert"] = TempTable;
  }

  // * Reset Temp table to prevent the data overriding
  TempTable = {};

  const ConvertionArrow = document.getElementById("ConvertArrow");

  // * Change the state of the convertion arrow
  ConvertionArrow.classList.remove("Disabled");
  ConvertionArrow.classList.add("Enable");
}

// + File Output Part
// > File conversion
async function FileConversion(json) {
  let header = await GetTableHeader();

  // * Place the first column header
  // & loop for each header in the output table
  const Tableheader = document.createElement("div");
  Tableheader.classList.add("TableParent", "row0");

  Tableheader.style.width = `${200 * 96}px`;

  // + String Input Handling
  function splitString(input) {
    const match = input.match(/^([A-Z]+)(\d+)$/);
    if (match) {
      const letters = match[1];
      const numbers = match[2];
      return [letters, numbers];
    }
    return null;
  }

  // * If the user is already selected the profile use for excel conversion
  if (CurrentProfile != null && Object.keys(FileDrawer).length > 0) {
    FileOutputMessage.style.display = "none";

    // * If the header of the output table is not existed, then create it. Otherwise, ignore.
    if (!OutputTempTable["row0"]) {
      // * Initialize the data
      OutputTempTable["row0"] = {};

      // * Loop till every header
      for (let i = 0; i < header.length; i++) {
        // * Create an element to hold the content of header of the table
        const TableHeaderElement = document.createElement("div");

        TableHeaderElement.classList.add("TableChild", `row0col${i}`);

        // * Set the limit for the characters
        TableHeaderElement.textContent = stringReduction(header[i], 15);

        // * A little styling to the table
        TableHeaderElement.style.borderTop = `rgb(230, 230, 230) solid 1px`;
        TableHeaderElement.style.width = "200px";

        // * Add the value into the output temp table
        OutputTempTable["row0"][`col${i}`] = header[i];

        // * Add the element to the table header
        Tableheader.appendChild(TableHeaderElement);
      }

      fileOutputContent.appendChild(Tableheader);
    }

    // * Assigning the current profile, both default and condition
    DefaultProfile = ProfileList[CurrentProfile.trim()].defaultProfile;
    ConditionalValue = ProfileList[CurrentProfile.trim()].conditionalValue;

    let AxisExist = false,
      count = 1;

    // * Determine whether the Y axis of the output table is existed or not
    if (document.querySelector(".YAxisIndicator.Output").innerHTML != "") {
      AxisExist = true;
    }

    let Max = 0;

    // * Find the maximum number of row
    for (let i = 0; i < Object.keys(FileDrawer).length; i++) {
      // * Get the data from the file drawer to display the result
      Data = FileDrawer[Object.keys(FileDrawer)[i]].Convert;

      Max += (findTimesToPrintColumn(Data) + 1) * 2;
    }

    // * Adding the file content according to the profile
    for (
      let i = 0;
      i < Object.keys(FileDrawer).length &&
      fileOutputContent.children.length < Max + 1;
      i++
    ) {
      // * Get the data from the file drawer to display the result
      Data = FileDrawer[Object.keys(FileDrawer)[i]].Convert;

      FileBackgroundColor =
        FileDrawer[Object.keys(FileDrawer)[i]]["BackgroundColor"];

      // * Find the loop count for the specific row (18 ~ TBD)
      let loop = findTimesToPrintColumn(Data);

      // * Generate the new row depend on the latest data
      // * Or generate the new axis depend the data
      if (
        count - fileOutputContent.children.length == 0 &&
        document.querySelector(".YAxisIndicator.Output").children.length < Max
      ) {
        if (
          document.querySelector(".YAxisIndicator.Output").children.length -
            1 +
            (loop + 1) * 2 ==
          Max
        ) {
          GenerateOutputAxis(Data);
        } else if (!AxisExist) {
          GenerateOutputAxis(Data);
        }
      }

      for (let k = 0; k < (loop + 1) * 2; k++, count++) {
        // * Create a parent content to hold the cell content
        const RowContent = document.createElement("div");

        RowContent.classList.add("TableParent", `row${count}`);
        RowContent.id = Object.keys(FileDrawer)[i].trim().split(".")[0];
        RowContent.style.width = "19200px";

        OutputTempTable[`row${count}`] = {};

        // * Loop to every profile
        DefaultProfile.forEach((dp, index) => {
          // * Create the `div` element
          const CellContent = document.createElement("div");
          CellContent.textContent = "";
          CellContent.classList.add("TableChild", `row${count}col${index}`);
          CellContent.style.width = `200px`;

          CellContent.style.backgroundColor = FileBackgroundColor;

          // * This is for TAX INVOICE  /  INVOICE / DEBIT NOTE
          if (index == 0 && k < loop + 1) {
            OutputTempTable[`row${count}`][`col${index}`] =
              "TAX INVOICE / INVOICE / DEBIT NOTE";
            CellContent.textContent = "TAX INVOICE / INVOICE / DEBIT NOTE";
          } else {
            // * In case it is not the conditional value
            if (dp && dp !== "CV") {
              const TargetCell = splitString(dp);

              // * Get the target Column
              let TargetColumn = TargetCell[0],
                TargetRow = TargetCell[1];

              // * Get the text from the extraction
              let Extrect = null

              console.log(TargetCell);

              if (Data[TargetRow][TargetColumn]) {
                Extrect = Data[TargetRow][TargetColumn];
              }

              // * Assign the output into the export table
              OutputTempTable[`row${count}`][`col${index}`] = Extrect;

              // * Reduce the length of the text
              Extrect = stringReduction(Extrect, 15);

              // * Substitute the text into the cell content
              CellContent.textContent = Extrect;
            }
            // * In case the conditional value
            else if (dp == "CV") {
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

                  let ConsiderationValue = "";

                  // * If condition consider about the output table
                  if (GivenCondition.split(" ")[1] == "OUTPUT") {
                    ConsiderationValue =
                      OutputTempTable[`row${count}`][
                        `col${getColumnIndex(GivenCondition.split(" ")[0]) - 1}`
                      ];

                    if (ConsiderationValue) {
                      ConsiderationValue = ConsiderationValue.trim().split(" ");
                    }
                  }

                  // * If the condition consider about the input part
                  else if (GivenCondition.split(" ")[1] == "INPUT") {
                    const TargetColumn = splitString(
                      GivenCondition.split(" ")[0]
                    );

                    ConsiderationValue = Data[TargetColumn[1]][TargetColumn[0]]
                      .trim()
                      .split(" ");
                  }

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
                    count,
                    index
                  );

                  break;
                case "OPERATION":
                  let OperationValue =
                    ConditionalValue[`${getColumnLabel(index)} column`][
                      ConditionHeader
                    ];

                  // * Length cutting algorithm
                  if (OperationValue[1] == "length") {
                    const TargetCell = splitString(OperationValue[0]);
                    let ConsiderationValue = Data[TargetCell[1]][TargetCell[0]];
                    let CutValue = "";

                    let Threshold = OperationValue[3];
                    Threshold = OperationValue[3];

                    if (Threshold < 0) {
                      Threshold =
                        ConditionalValue.length - Math.abs(OperationValue[3]);
                    }

                    switch (OperationValue[2]) {
                      // - If the value is `positive` : count from front | If the value is `negative` : count from back
                      // - This applies to every operation condition

                      case "start":
                        CutValue = ConsiderationValue.substring(
                          Threshold,
                          ConsiderationValue.length
                        );

                        break;

                      case "end":
                        CutValue = ConsiderationValue.substring(0, Threshold);

                        break;

                      // * This case is cut the string from `start` to `end`
                      case "from":
                        let endThreshold = OperationValue[4];

                        if (endThreshold < 0) {
                          endThreshold =
                            ConsiderationValue.length - Math.abs(endThreshold);
                        }

                        CutValue = ConsiderationValue.substring(
                          Threshold,
                          endThreshold
                        );

                        break;
                    }

                    OutputTempTable[`row${count}`][`col${index}`] = CutValue;

                    CutValue = stringReduction(CutValue, 15);

                    CellContent.textContent = CutValue;
                  } else {
                    let InitialString = "";

                    for (let i = 0; i < OperationValue.length; i++) {
                      if (OperationValue[i] == "add") continue;

                      let TargetCell = splitString(OperationValue[i]);
                      let AddingData = Data[TargetCell[1]][TargetCell[0]];

                      InitialString += AddingData;
                    }

                    OutputTempTable[`row${count}`][`col${index}`] =
                      InitialString;

                    // * Reduce the size of string
                    InitialString = stringReduction(InitialString, 15);

                    CellContent.textContent = InitialString;
                  }

                  break;
                case "DEFAULT":
                  let Text =
                    ConditionalValue[`${getColumnLabel(index)} column`][
                      ConditionHeader
                    ];

                  OutputTempTable[`row${count}`][`col${index}`] = Text;

                  // * Reduce the size of string
                  Text = stringReduction(Text, 15);

                  CellContent.textContent = Text;

                  break;
              }
            }

            // * In case it `null` value
            else {
              CellContent.textContent = " ";

              OutputTempTable[`row${count}`][`col${index}`] = null;
            }
          }

          // * Add the content to the row
          RowContent.appendChild(CellContent);
        });

        fileOutputContent.appendChild(RowContent);
      }
    }
  } else {
    if (Object.keys(FileDrawer) == 0) {
      FileOutputMessage.textContent = "Please input your excel file";
    } else {
      // * Warn the user to select the profile before converting
      FileOutputMessage.textContent = "Please select your profile";
    }

    FileOutputMessage.style.display = "flex";
  }
}

// > String reduction function
function stringReduction(text, length) {
  // * Check whether the string is longer than the threshold
  if (text && text.length > length) {
    return text.substring(0, length) + "...";
  }

  return text;
}

// > Convert from excel format `Date` to the actual `Date`
function excelDateToCustomFormat(serial) {
  // * Excel serial dates start from 1900-01-01
  const excelStartDate = new Date(1900, 0, 1);

  // * Add the serial number of days to the start date
  // * Note: Excel incorrectly considers 1900 as a leap year, so subtract 1 day
  const excelDate = new Date(
    excelStartDate.getTime() + (serial - 2) * 86400000
  );

  // * Options for formatting the date
  const options = { day: "numeric", month: "short", year: "2-digit" };

  // * Format the date
  return excelDate.toLocaleDateString("en-GB", options);
}

// + Finding the `times` to print out some column
// >
function findTimesToPrintColumn(Data) {
  // > Filter the rows where the length is between 16 and 43
  const rowsInRange = Object.keys(Data).filter((row) => row > 16 && row < 43);

  // > Return the count of such rows
  return rowsInRange.length;
}

// + Checking condition
// > Checking the value base on the condition
function ConditionConsideration(
  ConsiderationValue,
  ConsiderationCondition,
  CellContent,
  ConditionHeader,
  GivenCondition,
  Count,
  index
) {
  let ConditionChecking = true,
    count = 0;

  // * Condition Checking
  // & Condition is not about the null and not null
  if (
    !(
      ConsiderationCondition[0].trim().toLowerCase().includes("null") ||
      ConsiderationCondition[0].trim().toLowerCase().includes("not")
    )
  ) {
    ConsiderationCondition.forEach((cdc) => {
      // * Split the string if the string contains multiple words
      cdc
        .trim()
        .split(" ")
        .forEach((cdcs, index) => {
          // * Avoiding empty string comparing (JavaScript is dumb)
          if (ConsiderationValue[count] == "") {
            count++;
          }

          // * If the consideration condition value is "ELSE" and has not been checked yet. Then, trigger the condition.
          if (
            (cdcs.includes("ELSE") && cdcs.length == 4 && !ConditionChecking) ||
            (cdcs.includes(ConsiderationValue[count]) && !ConditionChecking)
          ) {
            ConditionChecking = true;
          } else if (!cdcs.includes(ConsiderationValue[count])) {
            ConditionChecking = false;
          }

          count++;
        });

      // * Initialize the count value
      count = 0;

      // * Reduce the string length of the `final value`
      if (ConditionChecking) {
        OutputTempTable[`row${Count}`][`col${index}`] = ConsiderationValue;

        ConsiderationValue = stringReduction(ConsiderationValue, 15);

        // * Assigning the value
        CellContent.textContent =
          ConditionalValue[`${getColumnLabel(index)} column`][ConditionHeader][
            GivenCondition
          ][cdc];
      }
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

      OutputTempTable[`row${Count}`][`col${index}`] = CellContent.textContent;
    } else {
      CellContent.textContent =
        ConditionalValue[`${getColumnLabel(index)} column`][ConditionHeader][
          GivenCondition
        ][ConsiderationCondition[0]];

      OutputTempTable[`row${count}`][`col${index}`] = CellContent.textContent;
    }
  }
}

// + Profile Button
// > Select profile for computing the result file
function SelectProfile(button) {
  let SelectedProfile = button.textContent;

  AddProfileText.textContent = SelectedProfile;
  CurrentProfile = SelectedProfile.trim();

  ProfileIndicator.textContent = SelectedProfile;
  ProfileIndicator.style.color = "black";

  fileOutputContent.innerHTML = "";

  // * Initialize the table again
  OutputTempTable = {};

  if (Object.keys(OutputTempTable).length == 0) {
    FileConversion(FileData);
  }

  DropdownContent();
}

// + Adding Excel Profile ------------
// > Openning Adding window profile
function OpenningAddingExcelProfile() {
  const ProfileAddingWindow = document.getElementById("AddingProfileWindow");
  ProfileAddingWindow.style.display = "flex";
}

// > Closing Adding window profile
function ClosingAddingExcelProfile() {
  const ProfileAddingWindow = document.getElementById("AddingProfileWindow");
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

// > Generate the input part axis
function GenerateInputAxis(json, highestNumber) {
  // % Variable declaration
  const XIndicatorInput = document.querySelector(".XAxisIndicator.Input");
  const YIndicatorInput = document.querySelector(".YAxisIndicator.Input");

  XIndicatorInput.innerHTML = "";
  YIndicatorInput.innerHTML = "";

  // % Generate the X axis Indicator - Input
  for (let i = 0; i <= highestNumber; i++) {
    // * Create `div` content
    const div = document.createElement("div");

    // * Adding classList
    div.classList.add("XIndicator");
    div.textContent = getColumnLabel(i);

    div.style.width = "150px";
    XIndicatorInput.appendChild(div);
  }

  // * Style the X indicator of input part
  XIndicatorInput.style.width = `${highestNumber * 150}px`;
  XIndicatorInput.style.paddingLeft = "30px";

  // % Generate the Y axis Indicator - Input
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

    // * Append content to the Y axis indicator
    YIndicatorInput.appendChild(div);
  });
}

// > Generate the output part axis
function GenerateOutputAxis(data) {
  // % Declare the variable of the output axis indicators
  const XIndicatorOutput = document.querySelector(".XAxisIndicator.Output");
  const YIndicatorOutput = document.querySelector(".YAxisIndicator.Output");

  // % Generate the X axis Indicator - Output
  // * Generate the X axis for the first time called, this doesn't change dynamically
  if (XIndicatorOutput.innerHTML == "") {
    for (let i = 0; i <= 96; i++) {
      const div = document.createElement("div");
      div.classList.add("XIndicator");
      div.textContent = getColumnLabel(i);
      div.style.width = "200px";
      XIndicatorOutput.appendChild(div);
    }

    XIndicatorOutput.style.width = `${96 * 200}px`;
    XIndicatorOutput.style.paddingLeft = "30px";
  }

  // % Generate the Y axis Indicator - Output
  // * If this function is called for the first time, then we need to add the header of the table
  // * Otherwise, we append the table with the remaining condition
  let loop = (findTimesToPrintColumn(data) + 1) * 2;

  CurrentChild = YIndicatorOutput.children.length;

  if (YIndicatorOutput.innerHTML == "") {
    loop += 1;
  }

  for (let i = CurrentChild; i < CurrentChild + loop; i++) {
    const div = document.createElement("div");
    div.classList.add("YIndicator");

    if (i == 0) {
      div.classList.add("first");
    }

    div.textContent = i + 1;
    YIndicatorOutput.appendChild(div);
  }
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

    DefaultProfile = data.defaultProfile;
    ConditionalValue = data.conditionalValue;
  } catch (err) {
    console.error("Error fetching profile:", err);
  }
}

// + Excel File Creation
// > Test function to download Excel
async function downloadExcel(filename = "data.xlsx") {
  if (Object.keys(OutputTempTable).length == 0) {
    await FileConversion(FileData);
  }

  const { ExportHeader, ExportData } = transformData(OutputTempTable);

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
      const row = jsonData[rowKey];

      let rowData = {};
      Object.keys(row).forEach((colKey, index) => {
        rowData[headers[index]] = row[colKey];
      });
      rows.push(rowData);
    });
  }

  return { ExportHeader: headers, ExportData: rows };
}

// + Warning Window Attributes
// > Closing the warning window function
function ClosingWarningWindow() {
  if (WarningCheck) {
    WarningCheck = false;
    WarningWindow.style.display = "none";
  } else {
    WarningCheck = true;
    WarningWindow.style.display = "block";
  }
}

// + Export Window Attribute
// > Export window attribute and functions
function TriggerExportWindow() {
  // * Get the export window
  const ExportWindow = document.getElementById("fileOutputWindow");

  // * Check whether the file is already convert or not
  if (CurrentProfile && CurrentFile) {
    // * Displaying and Closing the window trigger function
    if (WindowExport) {
      WindowExport = false;
      ExportWindow.style.display = "none";
    } else {
      WindowExport = true;
      ExportWindow.style.display = "block";
    }
  } else if (!CurrentFile) {
    WarningMessage.style.fontSize = `30px`;
    WarningMessage.textContent = "File content not found";

    ClosingWarningWindow();
  } else {
    WarningMessage.style.fontSize = `40px`;
    WarningMessage.textContent = "Invalid Profile";

    ClosingWarningWindow();
  }
}
