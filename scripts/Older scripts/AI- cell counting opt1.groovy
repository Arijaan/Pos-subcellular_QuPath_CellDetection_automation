/**
 * Automated miR-155 cell detection with parameter optimization
 * This script automatically finds optimal detection parameters and tests miR-155 thresholds
 */

import qupath.lib.common.GeneralTools
import static qupath.lib.gui.scripting.QPEx.*

// Define the output path
def pathOutput = buildFilePath(PROJECT_BASE_DIR, 'miR-155 optimized results')
mkdirs(pathOutput)

// Set H&E image type with default values
setImageType('BRIGHTFIELD_H_E');
clearAllObjects()

// Export the original image
def name = getProjectEntry().getImageName().replaceAll('.ome.tif', '')
def viewer = getCurrentViewer()
def path = buildFilePath(pathOutput, "$name original.png")
writeRenderedImage(viewer, path)

// Create an annotation around the full image
createSelectAllObject(true);

println "Starting parameter optimization for maximum cell detection..."

// Parameter ranges for optimization
def detectionThresholds = [0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4]
def sigmaMicrons = [0.5, 1.0, 1.5, 2.0, 2.5]
def minAreaMicrons = [5.0, 10.0, 15.0, 20.0]
def maxAreaMicrons = [200.0, 300.0, 400.0, 500.0]

def maxCellCount = 0
def optimalParams = [:]

// Find optimal detection parameters
for (threshold in detectionThresholds) {
    for (sigma in sigmaMicrons) {
        for (minArea in minAreaMicrons) {
            for (maxArea in maxAreaMicrons) {
                if (maxArea <= minArea) continue
                
                clearDetections()
                
                try {
                    // Test cell detection with current parameters
                    runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', 
                        '{"detectionImageBrightfield": "Optical density sum",  "requestedPixelSizeMicrons": 0.5,  "backgroundRadiusMicrons": 0.0,  "medianRadiusMicrons": 0.0,  "sigmaMicrons": ' + sigma + ',  "minAreaMicrons": ' + minArea + ',  "maxAreaMicrons": ' + maxArea + ',  "threshold": ' + threshold + ',  "maxBackground": 2.0,  "watershedPostProcess": true,  "excludeDAB": false,  "cellExpansionMicrons": 0.0,  "includeNuclei": true,  "smoothBoundaries": true,  "makeMeasurements": true,  "thresholdCompartment": "Nucleus: DAB OD mean",  "thresholdPositive1": 0.1,  "singleThreshold": true}');
                    
                    def cellCount = getDetectionObjects().size()
                    
                    if (cellCount > maxCellCount) {
                        maxCellCount = cellCount
                        optimalParams = [
                            threshold: threshold,
                            sigma: sigma,
                            minArea: minArea,
                            maxArea: maxArea
                        ]
                        println "New maximum: ${cellCount} cells with threshold=${threshold}, sigma=${sigma}, minArea=${minArea}, maxArea=${maxArea}"
                    }
                } catch (Exception e) {
                    // Skip problematic parameter combinations
                    continue
                }
            }
        }
    }
}

println "Optimization complete!"
println "Optimal parameters found: ${optimalParams}"
println "Maximum cell count: ${maxCellCount}"

// Now run miR-155 detection with optimal parameters and your specified thresholds
def results = [String.join('\t', ['Name', 'Detection threshold', 'miR-155 threshold', 'Num cells', 'Positive %', 'Negative %', '1+ %', '2+ %', '3+ %'])]

// miR-155 thresholds: 0.05, 0.10, 0.15 (with 3 levels each)
def miR155Thresholds = [0.05, 0.10, 0.15]

for (miR155Threshold in miR155Thresholds) {
    clearDetections()
    
    // Calculate the three threshold levels
    def threshold1 = miR155Threshold
    def threshold2 = miR155Threshold + 0.05
    def threshold3 = miR155Threshold + 0.10
    
    // Run detection with optimal parameters and miR-155 thresholds
    runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', 
        '{"detectionImageBrightfield": "Optical density sum",  "requestedPixelSizeMicrons": 0.5,  "backgroundRadiusMicrons": 0.0,  "medianRadiusMicrons": 0.0,  "sigmaMicrons": ' + optimalParams.sigma + ',  "minAreaMicrons": ' + optimalParams.minArea + ',  "maxAreaMicrons": ' + optimalParams.maxArea + ',  "threshold": ' + optimalParams.threshold + ',  "maxBackground": 2.0,  "watershedPostProcess": true,  "excludeDAB": false,  "cellExpansionMicrons": 0.0,  "includeNuclei": true,  "smoothBoundaries": true,  "makeMeasurements": true,  "thresholdCompartment": "Nucleus: DAB OD mean",  "thresholdPositive1": ' + threshold1 + ',  "thresholdPositive2": ' + threshold2 + ',  "thresholdPositive3": ' + threshold3 + ',  "singleThreshold": false}');

    // Write a rendered image
    path = buildFilePath(pathOutput, "$name miR-155=${GeneralTools.formatNumber(miR155Threshold, 2)}.png")
    writeRenderedImage(viewer, path)

    def percentages = calculateAllPercentages()
    
    results << String.join('\t', [
                name, 
                optimalParams.threshold as String,
                miR155Threshold as String,
                countCells() as String,
                percentages.positive as String,
                percentages.negative as String,
                percentages.onePlus as String,
                percentages.twoPlus as String,
                percentages.threePlus as String
                ])
}

// Print the results to the console
println '\n' + String.join('\n', results)

// Save the results in a tab-delimited file
def fileResults = new File(pathOutput, "$name-miR155-results.tsv")
fileResults.text = String.join('\n', results)

/**
 * Count the number of cells in total
 */
int countCells() {
    return getDetectionObjects().size()
}

/**
 * Calculate percentages for all classification levels
 */
Map calculateAllPercentages() {
    def detections = getDetectionObjects()
    def nCells = detections.size()
    
    if (nCells == 0) return [positive: 0, negative: 0, onePlus: 0, twoPlus: 0, threePlus: 0]
    
    def negative = detections.findAll {p -> p.getPathClass() == getPathClass('Negative')}.size()
    def onePlus = detections.findAll {p -> p.getPathClass() == getPathClass('1+')}.size()
    def twoPlus = detections.findAll {p -> p.getPathClass() == getPathClass('2+')}.size()
    def threePlus = detections.findAll {p -> p.getPathClass() == getPathClass('3+')}.size()
    def positive = onePlus + twoPlus + threePlus
    
    return [
        positive: positive * 100.0 / nCells,
        negative: negative * 100.0 / nCells,
        onePlus: onePlus * 100.0 / nCells,
        twoPlus: twoPlus * 100.0 / nCells,
        threePlus: threePlus * 100.0 / nCells
    ]
}