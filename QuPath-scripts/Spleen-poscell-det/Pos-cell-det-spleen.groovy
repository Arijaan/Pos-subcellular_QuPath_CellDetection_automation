// Create an annotation around the full image
clearAllObjects()
createSelectAllObject(true);

// Positive cell detection
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
