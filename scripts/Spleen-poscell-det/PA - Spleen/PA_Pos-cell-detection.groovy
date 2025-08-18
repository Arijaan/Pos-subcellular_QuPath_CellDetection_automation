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
            clearAllObjects()
            createSelectAllObject(true)
            // Run the same positive cell detection as 1-Final 3pos.groovy
            runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', [
            'detectionImageBrightfield': 'Optical density sum',
            'requestedPixelSizeMicrons': 0.7,
            'backgroundRadiusMicrons': 40.0,
            'backgroundByReconstruction': true,
            'medianRadiusMicrons': 0.8,
            'sigmaMicrons': 1,
            'minAreaMicrons': 5.0,
            'maxAreaMicrons': 400.0,
            'threshold': 0.03,
            'maxBackground': 0.7,
            'watershedPostProcess': true,
            'cellExpansionMicrons': 5.0,
            'includeNuclei': true,
            'smoothBoundaries': true,
            'makeMeasurements': true,
            'thresholdCompartment': 'Nucleus: miR-155 OD mean',
            'thresholdPositive1': 0.05,
            'thresholdPositive2': 0.1,
            'thresholdPositive3': 0.15,
            'singleThreshold': false
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


