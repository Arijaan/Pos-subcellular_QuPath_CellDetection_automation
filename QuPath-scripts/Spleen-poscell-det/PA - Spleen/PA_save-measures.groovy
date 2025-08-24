// Script to export measurements from all images in the project
// Uses only QuPath scripting API
// Applies the same processing as "Save measures.groovy" to each image

import qupath.lib.gui.tools.MeasurementExporter
import qupath.lib.objects.PathCellObject
import javafx.application.Platform
import java.util.concurrent.CountDownLatch

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

// Choose the columns that will be included in the export
def columnsToInclude = new String[]{"Image", "Classification", "Nucleus: miR-155 OD mean", "Nucleus: Hematoxylin OD mean"}

// Choose the type of objects that the export will process
def exportType = PathCellObject.class

for (imageEntry in imageEntries) {
    try {
        def imageName = imageEntry.getImageName()
        print "Processing: ${imageName}"
        
        // Read the image data
        def imageData = imageEntry.readImageData()
        if (imageData == null) {
            print "ERROR: Could not read image data for: ${imageName}"
            failedImages++
            failedImagesList.add(imageName + " (failed to read)")
            continue
        }
        
        // Set the current image data for processing using FX thread
        def latch = new CountDownLatch(1)
        def imageSetSuccessfully = false
        
        Platform.runLater {
            try {
                getCurrentViewer().setImageData(imageData)
                imageSetSuccessfully = true
            } catch (Exception e) {
                print "ERROR: Failed to set image in viewer: ${e.getMessage()}"
            } finally {
                latch.countDown()
            }
        }
        
        // Wait for the image to be set in the viewer
        latch.await()
        
        if (!imageSetSuccessfully) {
            print "ERROR: Could not set image in viewer for: ${imageName}"
            failedImages++
            failedImagesList.add(imageName + " (viewer failed)")
            continue
        }
        
        // Get the image name for the output filename
        def imageNameWithoutExtension = imageName.replaceAll("\\.[^.]*\$", "")
        // Clean filename by removing invalid characters for file system
        def cleanImageName = imageNameWithoutExtension.replaceAll("[\\\\/:*?\"<>|]", "_").trim()
        
        // Create output path for this specific image
        def path1 = "C:/Users/Bob/OneDrive/Desktop/Third QuPath/Measurements/${cleanImageName}_subcell_measurements.csv"
        def path2 = "C:/Users/psoor/OneDrive/Desktop/Measurements_ctrl/${cleanImageName}_subcell_measurements.csv"
        def outputPath = path2
        def outputFile = new File(outputPath)

        
        // Create the measurements folder if it doesn't exist
        def parentDir = outputFile.getParentFile()
        if (!parentDir.exists()) {
            def created = parentDir.mkdirs()
            if (!created) {
                print "ERROR: Could not create directory: ${parentDir.getAbsolutePath()}"
                failedImages++
                failedImagesList.add(imageName + " (directory creation failed)")
                continue
            }
        }
        
        // Create list with current image entry for export
        def imagesToExport = [imageEntry]
        
        // Create the measurementExporter and start the export
        def exporter = new MeasurementExporter()
                      .imageList(imagesToExport)            // Images from which measurements will be exported
                      .separator(separator)                 // Character that separates values
                      .includeOnlyColumns(columnsToInclude) // Columns are case-sensitive
                      .exportType(exportType)               // Type of objects to export
                      .exportMeasurements(outputFile)        // Start the export process
        
        // Add summary statistics to the same file
        def detections = imageData.getHierarchy().getDetectionObjects()
        def totalCells = detections.size()
        def negativeCells = detections.count { it.getPathClass()?.getName() == 'Negative' }
        def positive1Cells = detections.count { it.getPathClass()?.getName() == '1+' }
        def positive2Cells = detections.count { it.getPathClass()?.getName() == '2+' }
        def positive3Cells = detections.count { it.getPathClass()?.getName() == '3+' }
        def totalPositive = positive1Cells + positive2Cells + positive3Cells
        
        // Read the existing file and add summary columns
        if (outputFile.exists()) {
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
                
                // Write the modified content back to the file
                outputFile.text = modifiedLines.join('\n')
            }
        }
        
        print "  -> Exported ${totalCells} cells (${totalPositive} positive) to: ${outputFile.name}"
        processedImages++
        
    } catch (Exception e) {
        print "ERROR processing ${imageEntry.getImageName()}: ${e.getMessage()}"
        failedImages++
        failedImagesList.add(imageEntry.getImageName() + " (exception: " + e.getMessage() + ")")
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