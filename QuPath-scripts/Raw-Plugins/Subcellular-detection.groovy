removeDetections()
selectAnnotations()

runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
'detectionImageBrightfield': 'Optical density sum',
'requestedPixelSizeMicrons': 0.7,
'backgroundRadiusMicrons': 0,
'backgroundByReconstruction': true,
'medianRadiusMicrons': 0.8,
'sigmaMicrons': 1,
'minAreaMicrons': 5.0,
'maxAreaMicrons': 500.0,
'threshold': 0.25,
'maxBackground': 2.0,
'watershedPostProcess': true,
'cellExpansionMicrons': 8,
'includeNuclei': true,
'smoothBoundaries': true,
'makeMeasurements': true,
])

runPlugin('qupath.imagej.detect.cells.SubcellularDetection', [
'detection[miR-155]': 0.12,
'doSmoothing': true,
'splitByIntensity': true,
'splitByShape': false,
'spotSizeMicrons': 0.5,
'minSpotSizeMicrons': 0.5,
'maxSpotSizeMicrons': 3,
'includeClusters': true
])