import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.nio.file.StandardCopyOption
import java.io.File

// Get the current project
def project = getProject()
if (project == null) {
    print("No project is currently open!")
    return
}

// Get project directory
def projectDir = project.getPath().getParent()
print("Project directory: " + projectDir)

// Create a new folder called "copied_images" in the project directory
def copiedImagesDir = projectDir.resolve("copied_images")
if (!Files.exists(copiedImagesDir)) {
    Files.createDirectories(copiedImagesDir)
    print("Created directory: " + copiedImagesDir)
} else {
    print("Directory already exists: " + copiedImagesDir)
}

// Get all image entries in the project
def imageEntries = project.getImageList()
print("Found " + imageEntries.size() + " images in the project")

def successCount = 0
def failureCount = 0

// Loop through each image entry
for (imageEntry in imageEntries) {
    try {
        // Get the original image path from the server URIs
        def imageData = imageEntry.readImageData()
        def server = imageData.getServer()
        def uris = server.getURIs()
        
        if (uris.isEmpty()) {
            print("Warning: No URIs found for image: " + imageEntry.getImageName())
            failureCount++
            continue
        }
        
        // Get the first URI (primary image file)
        def uri = uris.iterator().next()
        print("Image URI: " + uri.toString())
        
        // Convert URI to file path
        def sourcePath = null
        try {
            if (uri.getScheme() == "file") {
                sourcePath = Paths.get(uri)
            } else {
                // Try to extract file path from URI string
                def uriString = uri.toString()
                if (uriString.startsWith("file:")) {
                    sourcePath = Paths.get(new java.net.URI(uriString))
                } else {
                    print("Warning: Cannot handle URI scheme for: " + uriString)
                    failureCount++
                    continue
                }
            }
        } catch (Exception uriEx) {
            print("Warning: Error parsing URI for " + imageEntry.getImageName() + ": " + uriEx.getMessage())
            failureCount++
            continue
        }
        
        if (sourcePath == null) {
            print("Warning: Could not determine source path for: " + imageEntry.getImageName())
            failureCount++
            continue
        }
        
        // Check if source file exists and is a file (not directory)
        if (!Files.exists(sourcePath)) {
            print("Warning: Source file does not exist: " + sourcePath)
            failureCount++
            continue
        }
        
        if (!Files.isRegularFile(sourcePath)) {
            print("Warning: Source is not a regular file: " + sourcePath)
            failureCount++
            continue
        }
        
        // Get the filename
        def fileName = sourcePath.getFileName().toString()
        
        // Create destination path
        def destinationPath = copiedImagesDir.resolve(fileName)
        
        // Check if file already exists at destination
        if (Files.exists(destinationPath)) {
            print("File already exists at destination, skipping: " + fileName)
            continue
        }
        
        // Copy the file
        print("Copying: " + fileName + " ...")
        Files.copy(sourcePath, destinationPath, StandardCopyOption.COPY_ATTRIBUTES)
        
        print("Successfully copied: " + fileName)
        successCount++
        
    } catch (Exception e) {
        print("Error copying " + imageEntry.getImageName() + ": " + e.getMessage())
        failureCount++
    }
}

print("\n=== Copy Summary ===")
print("Successfully copied: " + successCount + " files")
print("Failed to copy: " + failureCount + " files")
print("Files copied to: " + copiedImagesDir)