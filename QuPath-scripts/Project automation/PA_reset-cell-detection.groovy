// Script to open each image in the project once in the viewer

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
            imageEntry.saveImageData(imageData)
            print "  -> ${imageName} opened in viewer"
        } catch (Exception e) {
            print "  -> Error opening ${imageName}: ${e.getMessage()}"
        }
    }
    
    // Wait a moment before opening the next image
    //Thread.sleep(2000)
}

print "Finished deleting all objects!"
