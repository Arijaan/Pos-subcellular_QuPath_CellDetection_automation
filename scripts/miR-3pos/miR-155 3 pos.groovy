// Create an annotation around the full image
clearAllObjects()
createSelectAllObject(true);
// Cell detection - Three Threshold
runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', [
    'detectionImageBrightfield': 'Optical density sum',
    'requestedPixelSizeMicrons': 0.25,
    'backgroundRadiusMicrons': 30.0,
    'backgroundByReconstruction': true,
    'medianRadiusMicrons': 0.5,
    'sigmaMicrons': 1.5,
    'minAreaMicrons': 5.0,
    'maxAreaMicrons': 400.0,
    'threshold': 0.1,
    'maxBackground': 2.0,
    'watershedPostProcess': true,
    'cellExpansionMicrons': 7.0,
    'includeNuclei': true,
    'smoothBoundaries': true,
    'makeMeasurements': true,
    'thresholdCompartment': 'Nucleus: miR-155 OD mean',
    'thresholdPositive1': 0.05,
    'thresholdPositive2': 0.1,
    'thresholdPositive3': 0.15,
    'singleThreshold': false
])
