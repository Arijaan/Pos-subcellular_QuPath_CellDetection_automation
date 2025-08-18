clearAllObjects()
createSelectAllObject(true)

// cell detection (required for subcell)
runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
'detectionImageBrightfield': 'Optical density sum',
'requestedPixelSizeMicrons': 0.7,
'backgroundRadiusMicrons': 30.0,
'backgroundByReconstruction': true,
'medianRadiusMicrons': 1,
'sigmaMicrons': 1.2,
'minAreaMicrons': 5.0,
'maxAreaMicrons': 500.0,
'threshold': 0.05,
'maxBackground': 2.0,
'watershedPostProcess': true,
'cellExpansionMicrons': 4,
'includeNuclei': true,
'smoothBoundaries': true,
'makeMeasurements': true,
])

// Subcellular detection for miR-155
runPlugin('qupath.imagej.detect.cells.SubcellularDetection', [
'detection[miR-155]': 0.14,
'doSmoothing': true,
'splitByIntensity': true,
'splitByShape': false,
'spotSizeMicrons': 0.5,
'minSpotSizeMicrons': 0.5,
'maxSpotSizeMicrons': 5,
'includeClusters': true
])