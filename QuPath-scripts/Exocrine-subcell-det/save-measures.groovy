import qupath.lib.gui.tools.MeasurementExporter
import qupath.lib.objects.PathCellObject
import qupath.fx.dialogs.FileChoosers

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

// Outputdir selection
def outputDir = FileChoosers.promptForDirectory("Select output folder", null)
if (outputDir == null)
    return
// Choose output path with dynamic filename
def outputPath = outputDir.toPath().resolve("${cleanImageName}_subcell_measurements.csv").toString()
def outputFile = new File(outputPath)

// Separate each measurement value in the output file with a comma (",")
def separator = ","

// Choose the columns that will be included in the export
def columnsToInclude = new String[]{"Image", "Classification", "Nucleus: miR-155 OD mean", "Nucleus: Hematoxylin OD mean"}

// Choose the type of objects that the export will process
// Other possibilities include:
//    1. PathAnnotationObject
//    2. PathDetectionObject
//    3. PathRootObject

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

print "Exporting measurements to: ${outputFile.getAbsolutePath()}"

// Create the measurementExporter and start the export
def exporter  = new MeasurementExporter()
                  .imageList(imagesToExport)            // Images from which measurements will be exported
                  .separator(separator)                 // Character that separates values
                  .includeOnlyColumns(columnsToInclude) // Columns are case-sensitive
                  .exportType(exportType)               // Type of objects to export
                  .exportMeasurements(outputFile)        // Start the export process

// Add summary statistics to the same file
def detections = getCurrentImageData().getHierarchy().getDetectionObjects()
def totalCells = detections.size()
def negativeCells = detections.count { it.getPathClass()?.getName() == 'Negative' }
def positive1Cells = detections.count { it.getPathClass()?.getName() == '1+' }
def positive2Cells = detections.count { it.getPathClass()?.getName() == '2+' }
def positive3Cells = detections.count { it.getPathClass()?.getName() == '3+' }
def totalPositive = positive1Cells + positive2Cells + positive3Cells

// Read the existing file and add summary columns
def lines = outputFile.readLines()
def modifiedLines = []

// Add summary headers to the first line
if (lines.size() > 0) {
    modifiedLines.add(lines[0] + ",,,SUMMARY,")
    modifiedLines.add(lines.size() > 1 ? lines[1] + ",,,Statistic,Count" : ",,,Statistic,Count")
    
    // Add data rows with summary
    def summaryData = [
        "Num Cells,${totalCells}",
        "Num Positive,${totalPositive}",
        "Num 1+,${positive1Cells}",
        "Num 2+,${positive2Cells}",
        "Num 3+,${positive3Cells}",
        "Num Negative,${negativeCells}"
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

print "Done!"