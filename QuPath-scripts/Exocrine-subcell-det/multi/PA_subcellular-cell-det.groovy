import javafx.application.Platform

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

print "Opening ${imageEntries.size()} images sequentially..."

for (int i = 0; i < imageEntries.size(); i++) {
    def imageEntry = imageEntries[i]
    def imageName = imageEntry.getImageName()
    print "Opening image ${i + 1}/${imageEntries.size()}: ${imageName}"
    
    // Open the image in the viewer using FX thread
    Platform.runLater {
        try {
            def imageData = imageEntry.readImageData()
            getCurrentViewer().setImageData(imageData)
            // Create an annotation around the full image
            clearAllObjects()
            createSelectAllObject(true)
            // Watershed cell detection, complementing and required for subcellular cell detection
            runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
                'detectionImageBrightfield': 'Optical density sum',
                'requestedPixelSizeMicrons': 0.7,
                'backgroundRadiusMicrons': 30.0,
                'backgroundByReconstruction': true,
                'medianRadiusMicrons': 1,
                'sigmaMicrons': 1.2,
                'minAreaMicrons': 5.0,
                'maxAreaMicrons': 500.0,
                'threshold': 0.05,
                'maxBackground': 2.0,
                'watershedPostProcess': true,
                'cellExpansionMicrons': 4,
                'includeNuclei': true,
                'smoothBoundaries': true,
                'makeMeasurements': true,
            ])

            // Subcellular detection for miR-155
            runPlugin('qupath.imagej.detect.cells.SubcellularDetection', [
                'detection[miR-155]': 0.14,
                'doSmoothing': true,
                'splitByIntensity': true,
                'splitByShape': false,
                'spotSizeMicrons': 0.5,
                'minSpotSizeMicrons': 0.5,
                'maxSpotSizeMicrons': 3,
                'includeClusters': true
            ])
            
            imageEntry.saveImageData(imageData)
            
            print "  -> ${imageName} opened in viewer"
        } catch (Exception e) {
            print "  -> Error opening ${imageName}: ${e.getMessage()}"
        }
    }
    
    // Wait a moment before opening the next image
    //Thread.sleep(2000)
}

print "Finished opening all images!"