import qupath.lib.gui.tools.MeasurementExporter
import qupath.lib.objects.PathCellObject
import qupath.lib.objects.PathDetectionObject
import qupath.lib.gui.dialogs.Dialogs

// Script to save measurements after watershed cell detection and subcellular detection
// Follows the same structure as Save measures.groovy

// Get the current open image instead of all project images
def imageData = getCurrentImageData()
def project = getProject()
def entry = project.getEntry(imageData)
def imagesToExport = [entry]

// Get the image name for the output filename
def imageName = entry.getImageName()
def imageNameWithoutExtension = imageName.replaceAll("\\.[^.]*\$", "")
// Clean filename by removing invalid characters for file system
def cleanImageName = imageNameWithoutExtension.replaceAll("[\\\\/:*?\"<>|]", "_").trim()

// Define paths AFTER cleanImageName is created
def outputDir = Dialogs.promptForDirectory("Select output folder", null)
if (outputDir == null)
    return
// Choose output path with dynamic filename
def outputPath = outputDir.toPath().resolve("${cleanImageName}_subcell_measurements.csv").toString()
def outputFile = new File(outputPath)

// Separate each measurement value in the output file with a comma (",")
def separator = ","

// Choose the columns that will be included in the export for watershed + subcellular detection
def columnsToInclude = new String[]{
    "Image", 
    "Classification", 
    "Subcellular: miR-155: Num clusters",
    "Subcellular: miR-155: Num single spots"
}

// Export only cell objects
def exportType = PathCellObject.class

// Create the measurements folder if it doesn't exist
def parentDir = outputFile.getParentFile()
if (!parentDir.exists()) {
    def created = parentDir.mkdirs()
    if (!created) {
        print "ERROR: Could not create directory: ${parentDir.getAbsolutePath()}"
        return
    }
    print "Created directory: ${parentDir.getAbsolutePath()}"
}

// Saving the image data before outputing the data

entry.saveImageData(imageData)

print "Exporting watershed + subcellular measurements to: ${outputFile.getAbsolutePath()}"

// Get cell objects for counting
def hierarchy = getCurrentImageData().getHierarchy()
def cellObjects = hierarchy.getCellObjects()

// Create the measurementExporter and start the export
def exporter = new MeasurementExporter()
                  .imageList(imagesToExport)            // Images from which measurements will be exported
                  .separator(separator)                 // Character that separates values
                  .includeOnlyColumns(columnsToInclude) // Columns are case-sensitive
                  .exportType(exportType)               // Type of objects to export
                  .exportMeasurements(outputFile)        // Start the export process

// Add summary statistics for watershed + subcellular detection
// Cell-level statistics

// Read the existing file and add summary columns
def lines = outputFile.readLines()
def modifiedLines = []

// Create lists for each classification type
def negativeList = []
def positive1List = []
def positive2List = []
def positive3List = []

// Parse CSV to classify cells based on "Subcellular: miR-155: Num single spots" and clusters
if (lines.size() > 1) {
    def headerLine = lines[0].split(",")
    def spotsColumnIndex = -1
    def clustersColumnIndex = -1
    
    // Find the column indices for both spots and clusters
    for (int i = 0; i < headerLine.length; i++) {
        if (headerLine[i].trim() == "Subcellular: miR-155: Num single spots") {
            spotsColumnIndex = i
        }
        if (headerLine[i].trim() == "Subcellular: miR-155: Num clusters") {
            clustersColumnIndex = i
        }
    }
    
    // Process each data row (skip header)
    if (spotsColumnIndex != -1 && clustersColumnIndex != -1) {
        for (int rowIndex = 1; rowIndex < lines.size(); rowIndex++) {
            def dataRow = lines[rowIndex].split(",")
            if (dataRow.length > Math.max(spotsColumnIndex, clustersColumnIndex)) {
                try {
                    def spotsValue = dataRow[spotsColumnIndex].trim() as Double
                    def clustersValue = dataRow[clustersColumnIndex].trim() as Double
                    
                    // Combined classification logic (spots + clusters)
                    if (clustersValue == 0 && spotsValue == 0) {
                        negativeList.add(rowIndex)  // Negative: 0 clusters AND 0 spots
                    } else if (clustersValue == 1) {
                        if (spotsValue >= 2) {
                            positive2List.add(rowIndex)  // 2+: 1 cluster AND spots >= 2
                        } else {
                            positive1List.add(rowIndex)  // 1+: 1 cluster (regardless of spots if < 2)
                        }
                    } else if (clustersValue >= 2) {
                        positive3List.add(rowIndex)  // 3+: clusters >= 2
                    } else {
                        // No clusters (clustersValue == 0), classify by spots only
                        if (spotsValue == 0) {
                            negativeList.add(rowIndex)  // Negative: 0 spots
                        } else if (spotsValue >= 1 && spotsValue < 4) {
                            positive1List.add(rowIndex)  // 1+: 1-3 spots
                        } else if (spotsValue >= 4 && spotsValue < 10) {
                            positive2List.add(rowIndex)  // 2+: 4-9 spots
                        } else if (spotsValue >= 10) {
                            positive3List.add(rowIndex)  // 3+: >= 10 spots
                        }
                    }
                } catch (Exception e) {
                    // Skip rows with invalid data
                    print "Warning: Could not parse values in row ${rowIndex}: spots=${dataRow[spotsColumnIndex]}, clusters=${dataRow[clustersColumnIndex]}"
                }
            }
        }
    } else {
        print "Warning: Required columns not found in CSV. Spots index: ${spotsColumnIndex}, Clusters index: ${clustersColumnIndex}"
    }
}

// Cells and classifications
def totalCells = cellObjects.size()
def negativeCells = negativeList.size()
def positive1Cells = positive1List.size()
def positive2Cells = positive2List.size()
def positive3Cells = positive3List.size()
def totalPositiveCells = positive1Cells + positive2Cells + positive3Cells
def cellsWithSubcellular = positive1Cells + positive2Cells + positive3Cells  // Cells with spots > 0

// Add summary headers to the first line
if (lines.size() > 0) {
    modifiedLines.add(lines[0] + ",,,SUMMARY,")
    modifiedLines.add(lines.size() > 1 ? lines[1] + ",,,Statistic,Value" : ",,,Statistic,Value")
    
    // Add data rows with summary for watershed + subcellular
    def summaryData = [
        "Num Cells,${totalCells}",
        "Num Positive,${totalPositiveCells}",  
        "Num 1+,${positive1Cells}",
        "Num 2+,${positive2Cells}",
        "Num 3+,${positive3Cells}",
        "Num Negative,${negativeCells}",
    ]
    
    for (int i = 2; i < lines.size(); i++) {
        def summaryIndex = i - 2
        if (summaryIndex < summaryData.size()) {
            modifiedLines.add(lines[i] + ",,," + summaryData[summaryIndex])
        } else {
            modifiedLines.add(lines[i] + ",,,")
        }
    }
}

// Write the modified content back to the file
outputFile.text = modifiedLines.join('\n')

print "Done! Exported ${totalCells} cells with subcellular data and classification summary."
