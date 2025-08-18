// Create an annotation around the full image
clearAllObjects()
createSelectAllObject(true);

// set basal values for Haematoxylin and miR-155
setImageType('BRIGHTFIELD_H_DAB');
setColorDeconvolutionStains(
    '{"Name" : "H-DAB", ' +
    '"Stain 1" : "Hematoxylin", ' +
    '"Values 1" : "0.613 0.677 0.408", ' +
    '"Stain 2" : "miR-155", ' +
    '"Values 2" : "0.495 0.823 0.280", ' +
    '"Background" : "255 255 255"}'
);

runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', [
    'detectionImageBrightfield': 'Optical density sum',
    'requestedPixelSizeMicrons': 0.3,
    'backgroundRadiusMicrons': 30.0,
    'backgroundByReconstruction': true,
    'medianRadiusMicrons': 1.0,
    'sigmaMicrons': 1.2,
    'minAreaMicrons': 5.0,
    'maxAreaMicrons': 500.0,
    'threshold': 0.05,
    'maxBackground': 2.0,
    'watershedPostProcess': true,
    'cellExpansionMicrons': 6.924198250728864,
    'includeNuclei': true,
    'smoothBoundaries': true,
    'makeMeasurements': true
])
