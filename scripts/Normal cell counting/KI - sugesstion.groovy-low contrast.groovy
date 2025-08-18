
// Create an annotation around the full image
createSelectAllObject(true);

// Set up color deconvolution for Hematoxylin and miR-155 staining
setColorDeconvolutionStains('{"Name" : "Hematoxylin miR-155", "Stain 1" : "Hematoxylin", "Values 1" : "0.899 0.415 0.136", "Stain 2" : "miR-155", "Values 2" : "0.482 0.834 0.574", "Stain 3" : "Residual", "Values 3" : "0.002 -0.316 0.949", "Background" : " 193 185 204"}');

// Adjust display settings for maximum contrast
setChannelDisplayRange('Hematoxylin', 0, 0.5)
setChannelDisplayRange('miR-155', 0, 0.3)
setChannelDisplayRange('Residual', 0, 0.2)

// Extremely aggressive cell detection - detect everything possible
runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
    'detectionImageBrightfield': 'Hematoxylin OD',
    'requestedPixelSizeMicrons': 0.1457,
    'backgroundRadiusMicrons': 100.0,
    'backgroundByReconstruction': false,
    'medianRadiusMicrons': 0.0,
    'sigmaMicrons': 4.0,
    'minAreaMicrons': 1.0,
    'maxAreaMicrons': 2000.0,
    'threshold': 0.001,
    'maxBackground': 0.1,
    'watershedPostProcess': false,
    'cellExpansionMicrons': 0.5,
    'includeNuclei': true,
    'smoothBoundaries': false,
    'makeMeasurements': true
])
