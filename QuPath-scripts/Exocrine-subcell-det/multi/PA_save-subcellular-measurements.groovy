// Script to export measurements from all images in the project
// Uses only QuPath scripting API
// Applies the same processing as "Save subcellular measures.groovy" to each image

import qupath.lib.gui.tools.MeasurementExporter
import qupath.lib.objects.PathCellObject
import qupath.fx.dialogs.FileChoosers

def project = getProject()
if (project == null) {
    print 'No project is open!'
    return
}

def imageEntries = project.getImageList()
if (imageEntries.isEmpty()) {
    print 'No images found in project!'
    return
}

print "Exporting measurements from ${imageEntries.size()} images..."
print "=" * 60

def processedImages = 0
def failedImages = 0
def failedImagesList = []

// Separator for CSV export
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

for (imageEntry in imageEntries) {
    try {
        def imageName = imageEntry.getImageName()
        print "Processing: ${imageName}"
        
        // Set the current image data
        def imageData = imageEntry.readImageData()
        
        // Get the image name for the output filename
        def imageNameWithoutExtension = imageName.replaceAll("\\.[^.]*\$", "")
        // Clean filename by removing invalid characters for file system
        def cleanImageName = imageNameWithoutExtension.replaceAll("[\\\\/:*?\"<>|]", "_").trim()
        
        // Create output path for this specific image
        def outputDir = FileChoosers.promptForDirectory("Select output folder", null)
        if (outputDir == null)
            return
        // Choose output path with dynamic filename
        def outputPath = outputDir.toPath().resolve("${cleanImageName}_subcell_measurements.csv").toString()
        def outputFile = new File(outputPath)
        
        // Create the measurements folder if it doesn't exist
        def parentDir = outputFile.getParentFile()
        if (!parentDir.exists()) {
            def created = parentDir.mkdirs()
            if (!created) {
                print "ERROR: Could not create directory: ${parentDir.getAbsolutePath()}"
                continue
            }
        }
        
        // Get cell objects for counting
        def hierarchy = imageData.getHierarchy()
        def cellObjects = hierarchy.getCellObjects()
        
        if (cellObjects.isEmpty()) {
            print "  -> No cells found in ${imageName}, skipping..."
            continue
        }
        
        // Create the measurementExporter and start the export
        def exporter = new MeasurementExporter()
                          .imageList([imageEntry])              // Current image only
                          .separator(separator)                 // Character that separates values
                          .includeOnlyColumns(columnsToInclude) // Columns are case-sensitive
                          .exportType(exportType)               // Type of objects to export
                          .exportMeasurements(outputFile)        // Start the export process
        
        // Read the exported file and add summary columns with classification
        def lines = outputFile.readLines()
        def modifiedLines = []
        
        // Create lists for each classification type
        def negativeList = []
        def positive1List = []
        def positive2List = []
        def positive3List = []
        
        // Parse CSV to classify cells based on spots and clusters
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
                            print "    Warning: Could not parse values in row ${rowIndex}"
                        }
                    }
                }
            }
        }
        
        // Calculate classification counts
        def totalCells = cellObjects.size()
        def negativeCells = negativeList.size()
        def positive1Cells = positive1List.size()
        def positive2Cells = positive2List.size()
        def positive3Cells = positive3List.size()
        def totalPositiveCells = positive1Cells + positive2Cells + positive3Cells
        def cellsWithSubcellular = totalPositiveCells
        
        // Add summary headers to the first line
        if (lines.size() > 0) {
            print "    Adding summary data to CSV..."
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
            print "    Summary data prepared: ${summaryData.size()} items"
            
            for (int i = 2; i < lines.size(); i++) {
                def summaryIndex = i - 2
                if (summaryIndex < summaryData.size()) {
                    modifiedLines.add(lines[i] + ",,," + summaryData[summaryIndex])
                } else {
                    modifiedLines.add(lines[i] + ",,,")
                }
            }
            
            print "    Modified ${modifiedLines.size()} lines total"
        } else {
            print "    Warning: No lines found in exported CSV file"
        }
        
        // Write the modified content back to the file
        outputFile.text = modifiedLines.join('\n')
        
        print "  -> Exported ${totalCells} cells (${totalPositiveCells} positive) to: ${outputFile.name}"
        processedImages++
        
    } catch (Exception e) {
        print "ERROR processing ${imageEntry.getImageName()}: ${e.getMessage()}"
        failedImages++
        failedImagesList.add(imageEntry.getImageName() + " (${e.getMessage()})")
    }
    
    print ""
}

print "=" * 60
print "SUMMARY:"
print "Total images: ${imageEntries.size()}"
print "Successfully exported: ${processedImages}"
print "Failed: ${failedImages}"

if (!failedImagesList.isEmpty()) {
    print ""
    print "Failed images:"
    failedImagesList.each { imageName ->
        print "  - ${imageName}"
    }
}

print ""
print "Batch measurements export complete!"