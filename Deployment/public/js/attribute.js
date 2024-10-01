// + Variable declaration
let WarningCheck = false,
  DropdownState = false,
  fileDrawerdropdownState = false,
  IsOpeningManual = false,
  IsOpeningfileConfig = false;

let CurrentOpenfileConfig = null;

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

// + Showing the dropdown content
function FileDropdownContent() {
  if (fileDrawerdropdownState) {
    FileDrawerDropdown.style.display = "none";
    fileDrawerdropdownState = false;
  } else {
    FileDrawerDropdown.style.display = "block";
    fileDrawerdropdownState = true;
  }
}

// + Showing the `How to use` page
function OpenManual() {
  const ManualWindow = document.getElementById("ManualWindow");

  if (IsOpeningManual) {
    ManualWindow.style.display = "none";
    IsOpeningManual = false;
  } else {
    ManualWindow.style.display = "flex";
    IsOpeningManual = true;
  }
}

// + Showing the `file configuration` panel
function OpenFileConfigPanel(sheetName) {
  const fileConfigPanel = document.getElementById(
    `fileConfigPanel${sheetName}`
  );

  const activefileConfigPanel = document.getElementById(
    `fileConfigPanel${CurrentOpenfileConfig}`
  );

  if (activefileConfigPanel && CurrentOpenfileConfig != sheetName) {
    activefileConfigPanel.classList.remove(".Active");
    activefileConfigPanel.style.display = "none";

    IsOpeningfileConfig = false;
  }

  CurrentOpenfileConfig = sheetName;

  if (IsOpeningfileConfig) {
    fileConfigPanel.style.display = "none";
    fileConfigPanel.classList.remove("Active");
    IsOpeningfileConfig = false;
  } else {
    fileConfigPanel.style.display = "block";
    fileConfigPanel.classList.add("Active");
    IsOpeningfileConfig = true;
  }
}

// + Delete the file target file
function DeleteFile() {
  // % Get the keys from FileDrawer and iterate over them
  const fileKeys = Object.keys(FileDrawer);

  fileKeys.forEach((key, index) => {
    if (key == CurrentOpenfileConfig) {
      // % Remove the file from FileDrawer
      delete FileDrawer[CurrentOpenfileConfig];

      // % If the current open file is the same as the one being deleted
      if (CurrentOpenfileConfig == CurrentFile) {
        if (fileKeys.length === 1) {
          // % If it's only file left, reset the current file variables
          CurrentFile = null;
          CurrentOpenfileConfig = null;

          RearrangeDrawerDropdown(fileKeys[index - 1]);
        } else if (index === fileKeys.length - 1) {
          // % Display the current file
          DisplayTheFile(
            fileKeys[index - 1],
            FileDrawer[fileKeys[index - 1]].Json
          );

          FileConversion(FileDrawer[fileKeys[index - 1]].Json);

          // % Rearrange the file in the drawer
          RearrangeDrawerDropdown(fileKeys[index - 1]);
        } else {
          DisplayTheFile(
            fileKeys[index + 1],
            FileDrawer[fileKeys[index + 1]].Json
          );

          FileConversion(FileDrawer[fileKeys[index + 1]].Json);

          RearrangeDrawerDropdown(fileKeys[index + 1]);
        }
      } else {

        RearrangeDrawerDropdown(CurrentFile);
      }

      IsOpeningfileConfig = false;
      CurrentOpenfileConfig = null;
    }
  });
}

