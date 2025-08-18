import qupath.lib.projects.ProjectIO
import qupath.lib.images.servers.ImageServerProvider
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

// Get the current project
def project = getProject()
if (project == null) {
    print("No project is currently open!")
    return
}

// Get project directory and locate copied_images folder
def projectDir = project.getPath().getParent()
def copiedImagesDir = projectDir.resolve("copied_images")

print("Project directory: " + projectDir)
print("Looking for images in: " + copiedImagesDir)

// Check if copied_images folder exists
if (!Files.exists(copiedImagesDir)) {
    print("Error: copied_images folder does not exist at: " + copiedImagesDir)
    print("Please run the copy script first to create the copied_images folder.")
    return
}

// Get list of image files in the copied_images folder
def imageFiles = []
Files.list(copiedImagesDir).forEach { path ->
    if (Files.isRegularFile(path)) {
        def fileName = path.getFileName().toString().toLowerCase()
        // Check for common image file extensions
        if (fileName.endsWith('.tif') || fileName.endsWith('.tiff') || 
            fileName.endsWith('.jpg') || fileName.endsWith('.jpeg') || 
            fileName.endsWith('.png') || fileName.endsWith('.bmp') || 
            fileName.endsWith('.gif') || fileName.endsWith('.svs') || 
            fileName.endsWith('.ndpi') || fileName.endsWith('.vsi') ||
            fileName.endsWith('.czi') || fileName.endsWith('.lsm')) {
            imageFiles.add(path)
        }
    }
}

print("Found " + imageFiles.size() + " image files to load")

if (imageFiles.isEmpty()) {
    print("No image files found in the copied_images folder!")
    return
}

def successCount = 0
def failureCount = 0

// Process each image file
for (imagePath in imageFiles) {
    try {
        def fileName = imagePath.getFileName().toString()
        print("Processing: " + fileName)
        
        // Check if image with this name already exists in project
        def existingEntry = null
        for (entry in project.getImageList()) {
            if (entry.getImageName().equals(fileName)) {
                existingEntry = entry
                break
            }
        }
        
        if (existingEntry != null) {
            print("Image already exists in project, updating path: " + fileName)
            
            // Try to update the existing entry to point to the new location
            try {
                // Remove old entry and add new one with same name
                def imageData = existingEntry.readImageData()
                def hierarchy = imageData.getHierarchy()
                
                // Store the hierarchy data before removing
                def annotations = hierarchy.getAnnotationObjects()
                def detections = hierarchy.getDetectionObjects()
                
                // Create new entry from copied file
                def newEntry = project.addImage(imagePath.toString())
                newEntry.setImageName(fileName)
                
                // Load the new image data and restore hierarchy
                def newImageData = newEntry.readImageData()
                def newHierarchy = newImageData.getHierarchy()
                
                // Add back the annotations and detections
                newHierarchy.addPathObjects(annotations)
                newHierarchy.addPathObjects(detections)
                
                // Save the updated entry
                newEntry.saveImageData(newImageData)
                
                // Remove the old entry
                project.removeImage(existingEntry)
                
                print("Successfully updated: " + fileName)
                successCount++
                
            } catch (Exception updateEx) {
                print("Warning: Could not update existing entry, will skip: " + updateEx.getMessage())
                continue
            }
            
        } else {
            // Add new image to project
            print("Adding new image: " + fileName)
            def newEntry = project.addImage(imagePath.toString())
            newEntry.setImageName(fileName)
            
            print("Successfully added: " + fileName)
            successCount++
        }
        
    } catch (Exception e) {
        print("Error processing " + imagePath.getFileName() + ": " + e.getMessage())
        failureCount++
    }
}

// Save the project
try {
    project.syncChanges()
    print("Project changes synchronized")
} catch (Exception e) {
    print("Warning: Could not sync project changes: " + e.getMessage())
}

print("\n=== Load Summary ===")
print("Successfully processed: " + successCount + " images")
print("Failed to process: " + failureCount + " images")
print("Images loaded from: " + copiedImagesDir)
print("All images now reference files in the copied_images folder")

// Refresh the project in the UI
fireHierarchyUpdate()