// Create an annotation around the full image
createSelectAllObject(true);

// Cell detection
runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
    'detectionImageBrightfield': 'Optical density sum',
    'requestedPixelSizeMicrons': 0.5,
    'backgroundRadiusMicrons': 8.0,
    'backgroundByReconstruction': true,
    'medianRadiusMicrons': 0.0,
    'sigmaMicrons': 1.5,
    'minAreaMicrons': 10.0,
    'maxAreaMicrons': 1000.0,
    'threshold': 0.1,
    'maxBackground': 2.0,
    'watershedPostProcess': true,
    'cellExpansionMicrons': 5.0,
    'includeNuclei': true,
    'smoothBoundaries': true,
    'makeMeasurements': true
])

