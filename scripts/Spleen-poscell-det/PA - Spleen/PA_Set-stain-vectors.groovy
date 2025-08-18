/**
 * PA_Set-stain-vectors.groovy
 * 
 * Opens each image in the project sequentially, sets stain vectors, saves, and moves to next.
 */

import qupath.lib.images.ImageData
import qupath.lib.color.ColorDeconvolutionStains
import qupath.lib.gui.QuPathGUI
import javafx.application.Platform
import java.util.concurrent.CountDownLatch

def project = getProject()
if (project == null) {
    print "No project is open!"
    return
}

def imageEntries = project.getImageList()
print "Processing ${imageEntries.size()} images sequentially..."

// Define the stain vectors to apply
def stainJSON =  '{"Name" : "miR-155 spleen", ' +
    '"Stain 1" : "Hematoxylin", ' +
    '"Values 1" : "0.5212852210452145 0.7494100394680723 0.4082233592830086", ' +
    '"Stain 2" : "miR-155", ' +
    '"Values 2" : "0.3390198332403838 0.9190537662180316 0.20101175953190897", ' +
    '"Background" : "255 255 255"}'

def viewer = getCurrentViewer()

for (int i = 0; i < imageEntries.size(); i++) {
    def imageEntry = imageEntries[i]
    def imageName = imageEntry.getImageName()
    print "Opening image ${i + 1}/${imageEntries.size()}: ${imageName}"
    
    try {
        // Open the image in the viewer
        def imageData = imageEntry.readImageData()
        
        if (imageData == null) {
            print "  ERROR: Could not load image data for ${imageName}"
            continue
        }
        
        // Set the image in the viewer
        Platform.runLater {
            viewer.setImageData(imageData)
        }
        
        // Wait a moment for the image to load
        Thread.sleep(500)
        
        // Set image type to brightfield
        imageData.setImageType(ImageData.ImageType.BRIGHTFIELD_OTHER)
        
        // Set the color deconvolution stains
        def stains = ColorDeconvolutionStains.parseColorDeconvolutionStainsArg(stainJSON)
        imageData.setColorDeconvolutionStains(stains)
        
        // Save the changes
        imageEntry.saveImageData(imageData)
        
        print "  SUCCESS: Applied stain vectors and saved ${imageName}"
        
        // Brief pause before next image
        Thread.sleep(200)
        
    } catch (Exception e) {
        print "  ERROR: Failed to process ${imageName}: ${e.getMessage()}"
        e.printStackTrace()
    }
}

print "Finished processing all images."
