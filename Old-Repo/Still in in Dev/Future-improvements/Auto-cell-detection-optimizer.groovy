// Image-specific parameter selection based on image characteristics

clearAllObjects()
createSelectAllObject(true);

// Get image characteristics
def imageData = getCurrentImageData()
def server = imageData.getServer()
def imageName = server.getMetadata().getName()

// Analyze image properties to determine optimal parameters
def imageStats = analyzeImageCharacteristics()

// Select parameters based on image characteristics
def params = selectOptimalParameters(imageStats, imageName)

print("Selected parameters for image: ${imageName}")
print("Pixel size: ${params.pixelSize}")
print("Background radius: ${params.backgroundRadius}")
print("Threshold: ${params.threshold}")

// Apply selected parameters
runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', [
    'detectionImageBrightfield': 'Optical density sum',
    'requestedPixelSizeMicrons': params.pixelSize,
    'backgroundRadiusMicrons': params.backgroundRadius,
    'backgroundByReconstruction': true,
    'medianRadiusMicrons': params.medianRadius,
    'sigmaMicrons': params.sigma,
    'minAreaMicrons': params.minArea,
    'maxAreaMicrons': params.maxArea,
    'threshold': params.threshold,
    'maxBackground': 2.0,
    'watershedPostProcess': true,
    'cellExpansionMicrons': 6.924198250728864,
    'includeNuclei': true,
    'smoothBoundaries': true,
    'makeMeasurements': true,
    'thresholdCompartment': 'Nucleus: miR-155 OD mean',
    'thresholdPositive1': 0.05,
    'thresholdPositive2': 0.1,
    'thresholdPositive3': 0.15,
    'singleThreshold': false
])

def analyzeImageCharacteristics() {
    // Analyze image properties
    def server = getCurrentImageData().getServer()
    def pixelSize = server.getPixelCalibration().getAveragedPixelSizeMicrons()
    def imageType = getCurrentImageData().getImageType()
    
    // You can add more sophisticated analysis here
    return [
        pixelSize: pixelSize,
        imageType: imageType,
        brightness: "medium", // Could calculate actual brightness
        contrast: "normal"     // Could calculate actual contrast
    ]
}

def selectOptimalParameters(stats, imageName) {
    // Define parameter sets for different image types
    def parameterSets = [
        "low_contrast": [
            pixelSize: 0.25,
            backgroundRadius: 20.0,
            threshold: 0.03,
            minArea: 3.0,
            maxArea: 400.0,
            medianRadius: 1.0,
            sigma: 1.0
        ],
        "normal": [
            pixelSize: 0.3,
            backgroundRadius: 30.0,
            threshold: 0.05,
            minArea: 5.0,
            maxArea: 500.0,
            medianRadius: 1.5,
            sigma: 1.2
        ],
        "high_density": [
            pixelSize: 0.35,
            backgroundRadius: 40.0,
            threshold: 0.07,
            minArea: 7.0,
            maxArea: 600.0,
            medianRadius: 2.0,
            sigma: 1.5
        ]
    ]
    
    // Logic to determine which parameter set to use
    if (imageName.toLowerCase().contains("low") || imageName.toLowerCase().contains("contrast")) {
        return parameterSets["low_contrast"]
    } else if (imageName.toLowerCase().contains("dense") || imageName.toLowerCase().contains("high")) {
        return parameterSets["high_density"]
    } else {
        return parameterSets["normal"]
    }
}
