// Clear all existing objects and create an annotation around the full image
clearAllObjects()
createSelectAllObject(true);

// cell detection
runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
    'detectionImageBrightfield': 'Optical density sum',
    'requestedPixelSizeMicrons': 0.5,
    'backgroundRadiusMicrons': 40.0,
    'backgroundByReconstruction': true,
    'medianRadiusMicrons': 0.5,
    'sigmaMicrons': 1.5,
    'minAreaMicrons': 10.0,
    'maxAreaMicrons': 1000.0,
    'threshold': 0.1,
    'maxBackground': 1.0,
    'watershedPostProcess': true,
    'cellExpansionMicrons': 9.0,
    'includeNuclei': true,
    'smoothBoundaries': true,
    'makeMeasurements': true
])