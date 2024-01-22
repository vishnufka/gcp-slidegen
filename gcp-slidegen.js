// ADD NEW PERMANENT FOLDERS HERE
const DEFAULT_FOLDER_COLUMNS = ["News Sites", "Product", "Training"]

// Slide Generator/slides folder in top-level-folder
const TOP_LEVEL_SLIDE_FOLDER_ID = "folder-id-go-here"  // TEST folder
//const TOP_LEVEL_SLIDE_FOLDER_ID = "" // !!!PROD folder
// Customers folder in top-level-folder
const TOP_LEVEL_CUSTOMER_FOLDER_ID = "folder-id-go-here" // TEST folder
//const TOP_LEVEL_CUSTOMER_FOLDER_ID = "" // !!!PROD folder

// hard-coded special folder names
const EXEC_FOLDER_NAME = "exec"
const HEADER_FOLDER_NAME = "header"
const CHECKINS_FOLDER_NAME = "checkins"
const CONSULTANT_FOLDER_NAME = "consultant"
const CUSTOMER_FOLDER_NAME = "customer"

// excel headers
const SLIDES_GENERATED_COL = "Last Generated On"
const CUSTOMER_NAME_COL = "Customer Name"
const LAST_MEETING_COL = "Last Meeting"
const NEXT_MEETING_COL = "Next Meeting"
const CONSULTANT_FOLDER_NAME_COL = "Custom Consultant Folders"

/** 
 * Creates the menu item "Generate Slides" inside Google Sheets for user to execute program.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Slides')
      .addItem('Generate Slides', 'generateSlides')
      .addToUi();
}
 
 /** 
 * MAIN PROGRAM ENTRY POINT. Modified version of the mail merge email code found here:
 * https://developers.google.com/apps-script/samples/automations/mail-merge
 * Executed by the user using the Google Sheets drop down.
 */
function generateSlides(customerName, sheet=SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!customerName){
    customerName = Browser.inputBox("Generate Slides", 
                                      "Type the name of the customer to generate slides for",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (customerName === "cancel" || customerName == ""){ 
    // If no subject line, finishes up
    return;
    }
  }
  
  // Gets the data from the passed sheet - took this from the mail merge sheet and modified
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 
  
  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const slidesGeneratedColIdx = heads.indexOf(SLIDES_GENERATED_COL);
  
  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record generated slides
  var out = [];

  var url = ""
  var clientRow = 0

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){

    // Only generate slides if SLIDES_GENERATED_COL cell is blank and not hidden by a filter
    if (row[SLIDES_GENERATED_COL] == '' && row[CUSTOMER_NAME_COL].toLowerCase() == customerName.toLowerCase()){
      try {
        clientRow = rowIdx
        url = slideGenerator(row, customerName)
        var today = Utilities.formatDate(new Date(), "GMT+0", 'yyyy-MM-dd')
        out.push([SpreadsheetApp.newRichTextValue().setText(today).setLinkUrl(url).build()])
      } catch(e) {
        // modify cell to record error
        out.push([SpreadsheetApp.newRichTextValue().setText(e.message).build()]);
      }
      sheet.getRange(clientRow+2, slidesGeneratedColIdx+1, 1).setRichTextValues(out)
    } else {
      // do nothing
    }
  });

}


 /** 
 * Slide generation logic fuction.
 */
function slideGenerator(row, customerName) {
  Logger.log("Starting slide generator for: " + customerName)
  var allLevelsFoldersDict = getAllLevelsFoldersDict()

  // get dates, if incorrect then abort else print
  var lastMeetingDateStr = row[LAST_MEETING_COL]
  var nextMeetingDateStr = row[NEXT_MEETING_COL]
  var lastMeetingDate = new Date(lastMeetingDateStr)
  var nextMeetingDate = new Date(nextMeetingDateStr)
  if (!isValidDate(lastMeetingDate) || !isValidDate(nextMeetingDate)) {
    var errorMsg = "ERROR: next and/or last meeting date not in YYYY-MM-DD format"
    Logger.log(errorMsg)
    throw new Error(errorMsg)
  }
  else {
    Logger.log("Last Meeting Date: " + lastMeetingDateStr)
    Logger.log("Next Meeting Date: " + nextMeetingDateStr)
  }

  // get folders to search, if empty then abort else print
  var defaultFoldersToSearchArr = getDefaultFoldersToSearchArr(row)
  var consultantFoldersToSearchArr = getConsultantFoldersToSearchArr(row)
  var foldersToSearchArr = consultantFoldersToSearchArr.concat(defaultFoldersToSearchArr)
  if (foldersToSearchArr.length == 0) {
    var errorMsg = "ERROR: no folders to search defined either with 1 2 3 etc or custom consultant folders"
    Logger.log(errorMsg)
    throw new Error(errorMsg)
  }
  else {
    // forces addition of exec folder to the start of the array
    foldersToSearchArr.unshift(EXEC_FOLDER_NAME)
    Logger.log("Folders to search in order: " + foldersToSearchArr)
  }

  // gets outputFolder
  var outputFolder = getOutputFolder(row[CUSTOMER_NAME_COL])
  var outputFilename = nextMeetingDateStr + " " + row[CUSTOMER_NAME_COL] + " Check-In"
  var filesInOutputFolder = DriveApp.getFolderById(outputFolder).getFiles()
  while (filesInOutputFolder.hasNext()) {
    var nextFile = filesInOutputFolder.next()
    if (nextFile.getName() == outputFilename) {
      var errorMsg = "ERROR: file named '" + outputFilename + "' already exists in folder '" + row[CUSTOMER_NAME_COL] + '/' + CHECKINS_FOLDER_NAME + "'"
      Logger.log(errorMsg)
      throw new Error(errorMsg)
    }
  }

  try {
    var outputPres = SlidesApp.create(outputFilename)

    // remove the blank first slide
    outputPres.getSlides().pop().remove()

    // adds customer slide
    Logger.log("Adding Customer slide")
    addCustomerSlides(outputPres, allLevelsFoldersDict, customerName.toLowerCase(), nextMeetingDateStr)
    Logger.log("DONE Adding Customer slide")

    // adds in-date slides for each folder
    var totalFolders = foldersToSearchArr.length
    foldersToSearchArr.forEach(function(folder, i) {
      Logger.log("Adding slides for folder " + (i+1) + " of " + totalFolders + " - " + folder)
      addFolderSlides(outputPres, allLevelsFoldersDict, folder, lastMeetingDate)
      Logger.log("DONE Adding slides for folder " + folder)
    })

    DriveApp.getFileById(outputPres.getId()).moveTo(DriveApp.getFolderById(outputFolder))
    Logger.log("Saved presentation to: " + outputPres.getUrl())
    Logger.log("SUCCESS: Finished slide generator for: " + customerName)

    return outputPres.getUrl()

  }
  // if the script fails we delete the half-finished deck
  catch (error) {
    DriveApp.getFileById(outputPres.getId()).setTrashed(true)
    Logger.log("FAILURE: Script aborted due to error")
    throw error
  }

}

 /** 
 * Returns dictionary of folder names and IDs, including special consultant folders
 */
function getAllLevelsFoldersDict() {
  
  var topLevelFoldersDict = getFoldersDict(TOP_LEVEL_SLIDE_FOLDER_ID)
  var consultantFoldersDictNames = getFoldersDict(topLevelFoldersDict[CONSULTANT_FOLDER_NAME])

  var consultantFoldersDict = Object.fromEntries(
    Object.entries(consultantFoldersDictNames).map(([key, value]) => 
      [CONSULTANT_FOLDER_NAME+ "/" + `${key}`.toLowerCase(), value]
    )
  )

  var allLevelsFoldersDict = Object.assign({}, topLevelFoldersDict, consultantFoldersDict)

  return allLevelsFoldersDict
}

/**
 * Adds customer template slides as first slides
 */
function addCustomerSlides(outputPres, allLevelsFoldersDict, customerName, nextMeetingDate) {
  var customerInitialFolderId = getOutputFolderFromInitial(customerName, allLevelsFoldersDict[CUSTOMER_FOLDER_NAME])
  var customerSlideFiles = DriveApp.getFolderById(customerInitialFolderId).getFiles()
  while (customerSlideFiles.hasNext()) {
    var nextCustomerSlideFile = customerSlideFiles.next()
    if (nextCustomerSlideFile.getName().toLowerCase() == customerName.toLowerCase()) {
      var customerSlides = SlidesApp.openById(nextCustomerSlideFile.getId()).getSlides()
      var len = outputPres.getSlides().length // will be 0
      for (var i = 0; i < customerSlides.length; i++) {
        outputPres.insertSlide(len++, customerSlides[i]).replaceAllText("INSERTDATE", nextMeetingDate)
      }
    }
  }
}

 /** 
 * Returns array of folder IDs to search, from the default folders offered.
 */
function getDefaultFoldersToSearchArr(row) {
  // the below statement can be removed eventually,
  // just need to deal with anyone still using the beta sheet
  if (row["FieldTeam"]) {
      var errorMsg = "ERROR: FieldTeam is deprecated (use Product instead), delete the FieldTeam column from this sheet"
      Logger.log(errorMsg)
      throw new Error(errorMsg)
  }

  var defaultFoldersToSearchDict = {}

  DEFAULT_FOLDER_COLUMNS.forEach(function(col, i) {
    if (row[col]) {
      defaultFoldersToSearchDict[row[col]] = col.toLowerCase()
    }
  })

  var defaultFoldersToSearchArr = []
  for (let i = 1; i <= Object.keys(defaultFoldersToSearchDict).length; i++) {
    defaultFoldersToSearchArr.push(defaultFoldersToSearchDict[i])
  }

  if (defaultFoldersToSearchArr.includes(undefined)) {
      var errorMsg = "ERROR: undefined element found, check you have used 1 2 3 etc in the correct order"
      Logger.log(errorMsg)
      throw new Error(errorMsg)
  }

  return defaultFoldersToSearchArr
}

 /** 
 * Returns array of folder IDs to search, from the consultant folders offered.
 * Split up the semicolon separated folders and return as array of consultant/foldername
 */
function getConsultantFoldersToSearchArr(row) {
  var consultantFoldersToSearchArr = []
  if (row[CONSULTANT_FOLDER_NAME_COL]) {
    consultantFoldersToSearchArr = row[CONSULTANT_FOLDER_NAME_COL].split(";").filter(el => el).map(folder => {
      return CONSULTANT_FOLDER_NAME + "/" + folder.trim().toLowerCase()
    })
  }

  return consultantFoldersToSearchArr
}


 /** 
 * Returns folder in Customers folder where the OutputPres will be saved.
 */
function getOutputFolder(customerName) {

  var customerInitialFolder = getOutputFolderFromInitial(customerName, TOP_LEVEL_CUSTOMER_FOLDER_ID)

  // find the customer name folder within the initial folder
  var customerNameFolderDict = getFoldersDict(customerInitialFolder)
  if (customerName in customerNameFolderDict) {
    var customerNameFolder = customerNameFolderDict[customerName]
  }
  else {
    var errorMsg = "ERROR: no customer folder named '" + customerName + "'"
    Logger.log(errorMsg)
    throw new Error(errorMsg)
  }

  // find the checkins folder within the customer name folder
  var checkinsFolderDict = getFoldersDict(customerNameFolder)
  if (CHECKINS_FOLDER_NAME in checkinsFolderDict) {
    var checkinsFolder = checkinsFolderDict[CHECKINS_FOLDER_NAME]
  }
  else {
    var errorMsg = "ERROR: no folder named '" + CHECKINS_FOLDER_NAME + "' (all lowercase, no dash) inside customer folder named '" + customerName + "'"
    Logger.log(errorMsg)
    throw new Error(errorMsg)
  }

  var outputFolder = checkinsFolder
  return outputFolder
}

 /** 
 * Logic to add slides to the OutputPres from a folder of slides.
 */
function addFolderSlides(outputPres, allLevelsFoldersDict, folder, lastMeetingDate) {
  try {
    var filesInFolder = DriveApp.getFolderById(allLevelsFoldersDict[folder]).getFiles()
  }
  catch {
    var errorMsg = "ERROR: Folder '" + folder + "' does not exist in top level folder"
    Logger.log(errorMsg)
    throw new Error(errorMsg)
  }
  
  // searches files in folder and compares to last meeting date
  var filesToAddArr = []
  while (filesInFolder.hasNext()) {
    var nextFile = filesInFolder.next()
    var fileName = nextFile.getName()
    var fileDate = new Date(fileName.substring(0,10))
    if (isValidDate(fileDate)) {
      if (fileDate >= lastMeetingDate) {
        filesToAddArr.push(nextFile)
      }
    }
    else {
      // not aborting here, as end users may not be able to edit the folders where slides are stored
      Logger.log("ERROR: File '" + fileName + "' does not start with YYYY-MM-DD fomatted date")
    }
  }

  // if no slides in date, continue, else add them
  if (filesToAddArr.length > 0) {
    Logger.log("Found " + filesToAddArr.length + " in-date files in " + folder)
    // adds header slide for that folder
    var headerSlideFiles = DriveApp.getFolderById(allLevelsFoldersDict[HEADER_FOLDER_NAME]).getFiles()
    while (headerSlideFiles.hasNext()) {
      var nextHeaderSlideFile = headerSlideFiles.next()
      if (nextHeaderSlideFile.getName().toLowerCase() == folder.toLowerCase()) {
        Logger.log("Adding header slide for " + folder)
        var headerSlide = SlidesApp.openById(nextHeaderSlideFile.getId()).getSlides()[0]
        outputPres.insertSlide(outputPres.getSlides().length, headerSlide)
        Logger.log("DONE Adding header slide for " + folder)
      }
    }

    // adds slides from in-date slide files
    filesToAddArr.sort().forEach(function(file, i) {
      Logger.log("Adding file " + (i+1) + " of " + filesToAddArr.length + " - " + file)
      var fileSlides = SlidesApp.openById(file.getId()).getSlides()
      var len = outputPres.getSlides().length
      for (var j = 0; j < fileSlides.length; j++) {
        outputPres.insertSlide(len++, fileSlides[j])
      }
      Logger.log("DONE Adding file " + file)
    })
  }
  else {
    Logger.log("No in-date files found in " + folder + " - continuing")
  }
}

 /** 
 * 
 * Helper functions
 * 
 */

 /** 
 * Returns dictionary of folder names and IDs from a top-level folder ID.
 */
function getFoldersDict(folderId) {
  var foldersDict = {}
  var foldersInFolder = DriveApp.getFolderById(folderId).getFolders()

  while (foldersInFolder.hasNext()) {
    var nextFolder = foldersInFolder.next()
    foldersDict[nextFolder.getName()] = nextFolder.getId()
  }

  return foldersDict
}

 /** 
 * Returns folder based on first character of customer name.
 */
function getOutputFolderFromInitial(customerName, topLevelFolder) {

  // find the initial folder, A B C D etc.
  var customerInitialFoldersDict = getFoldersDict(topLevelFolder)
  var customerInitial = customerName.substring(0,1).toUpperCase()
  // if it starts with a number we need to get the 123 folder
  if (!isNaN(customerInitial)) {
    customerInitial = '123'
  }
  var customerInitialFolder = customerInitialFoldersDict[customerInitial]
  return customerInitialFolder
}

 /** 
 * Checks string is a valid date
 */
function isValidDate(d) {
  return (d instanceof Date && !isNaN(d))
}
