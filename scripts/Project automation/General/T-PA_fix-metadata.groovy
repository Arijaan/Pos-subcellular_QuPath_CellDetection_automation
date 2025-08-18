import qupath.lib.images.servers.PixelCalibration
import qupath.lib.images.ImageData
import qupath.lib.display.ChannelDisplayInfo
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

// Go through every image and set pixel calibration
def fixedCount = 0
def viewer = getCurrentViewer()

for (int i = 0; i < imageEntries.size(); i++) {
    def imageEntry = imageEntries[i]
    def imageName = imageEntry.getImageName()
    print "Processing image ${i + 1}/${imageEntries.size()}: ${imageName}"
    
    try {
        // Read the image data
        def imageData = imageEntry.readImageData()
        if (imageData == null) {
            print "  ERROR: Could not load image data for: ${imageName}"
            continue
        }
        
        // Set the image in the viewer using FX thread with synchronization
        def latch = new CountDownLatch(1)
        Platform.runLater {
            try {
                viewer.setImageData(imageData)
                latch.countDown()
            } catch (Exception e) {
                print "  ERROR: Failed to set image in viewer: ${e.getMessage()}"
                latch.countDown()
            }
        }
        
        // Wait for the image to be set in the viewer
        latch.await()
        Thread.sleep(500)
        
        // Set pixel calibration to static values
        setPixelSizeMicrons(0.145600, 0.145600)
        
        // Get the updated image data from the viewer and save it
        def updatedImageData = viewer.getImageData()
        imageEntry.saveImageData(updatedImageData)
        
        fixedCount++
        print "  SUCCESS: Fixed pixel calibration for ${imageName}"
        
    } catch (Exception e) {
        print "  ERROR: Failed to process ${imageName}: ${e.getMessage()}"
        e.printStackTrace()
    }
}

print "Pixel calibration fix complete! Fixed ${fixedCount} images."