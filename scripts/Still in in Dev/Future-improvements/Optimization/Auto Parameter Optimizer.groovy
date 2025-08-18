// Automated Parameter Optimizer for Positive Cell Detection

// Clear existing objects
clearAllObjects()
createSelectAllObject(true);

// Define parameter ranges to test
def thresholds = [0.03, 0.05, 0.07]
def medianRadii = [0.5, 1.0, 1.5]
def sigmas = [1.0, 1.2, 1.5]

// Fixed parameters
def fixedPixelSize = 0.3
def fixedBackgroundRadius = 30.0
def fixedMinArea = 5.0
def fixedMaxArea = 500.0

def bestParams = null
def bestScore = 0
def allResults = []

// Test different parameter combinations
for (threshold in thresholds) {
    for (medianRadius in medianRadii) {
        for (sigma in sigmas) {
            
            // Clear previous detection
            clearAllObjects()
            createSelectAllObject(true);
            
            // Run detection with current parameters
            try {
                runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', [
                    'detectionImageBrightfield': 'Optical density sum',
                    'requestedPixelSizeMicrons': fixedPixelSize,
                    'backgroundRadiusMicrons': fixedBackgroundRadius,
                    'backgroundByReconstruction': true,
                    'medianRadiusMicrons': medianRadius,
                    'sigmaMicrons': sigma,
                    'minAreaMicrons': fixedMinArea,
                    'maxAreaMicrons': fixedMaxArea,
                    'threshold': threshold,
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
                
                // Evaluate detection quality
                def detections = getDetectionObjects()
                def cellCount = detections.size()
                
                if (cellCount > 0) {
                    // Calculate quality metrics
                    def avgNucleusArea = detections.collect { it.getNucleusROI()?.getArea() ?: 0 }.sum() / cellCount
                    def avgOD = detections.collect { it.getMeasurementList().getMeasurementValue('Nucleus: miR-155 OD mean') }.sum() / cellCount
                    
                    // Calculate miR-155 stain evaluation metrics
                    def stainMetrics = evaluateMiR155Stain(detections)
                    
                    // Calculate score (customize based on your criteria)
                    def score = calculateQualityScore(cellCount, avgNucleusArea, avgOD, stainMetrics)
                    
                    def result = [
                        threshold: threshold,
                        medianRadius: medianRadius,
                        sigma: sigma,
                        cellCount: cellCount,
                        avgNucleusArea: avgNucleusArea,
                        avgOD: avgOD,
                        negativePercent: stainMetrics.negativePercent,
                        positive1Percent: stainMetrics.positive1Percent,
                        positive2Percent: stainMetrics.positive2Percent,
                        positive3Percent: stainMetrics.positive3Percent,
                        stainDistribution: stainMetrics.stainDistribution,
                        score: score
                    ]
                    
                    allResults.add(result)
                    
                    if (score > bestScore) {
                        bestScore = score
                        bestParams = result
                    }
                    
                    print("Tested: thresh=${threshold}, median=${medianRadius}, sigma=${sigma}, cells=${cellCount}, neg=${stainMetrics.negativePercent}%, 1+=${stainMetrics.positive1Percent}%, score=${score}")
                }
                
            } catch (Exception e) {
                print("Failed with parameters: thresh=${threshold}, median=${medianRadius}, sigma=${sigma}")
            }
        }
    }
}

// Apply best parameters
if (bestParams != null) {
    clearAllObjects()
    createSelectAllObject(true);
    
    runPlugin('qupath.imagej.detect.cells.PositiveCellDetection', [
        'detectionImageBrightfield': 'Optical density sum',
        'requestedPixelSizeMicrons': fixedPixelSize,
        'backgroundRadiusMicrons': fixedBackgroundRadius,
        'backgroundByReconstruction': true,
        'medianRadiusMicrons': bestParams.medianRadius,
        'sigmaMicrons': bestParams.sigma,
        'minAreaMicrons': fixedMinArea,
        'maxAreaMicrons': fixedMaxArea,
        'threshold': bestParams.threshold,
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
    
    print("BEST PARAMETERS APPLIED:")
    print("Threshold: ${bestParams.threshold}")
    print("Median Radius: ${bestParams.medianRadius}")
    print("Sigma: ${bestParams.sigma}")
    print("Score: ${bestParams.score}")
    print("Cell Count: ${bestParams.cellCount}")
    print("Negative: ${bestParams.negativePercent}%")
    print("1+: ${bestParams.positive1Percent}%")
    print("2+: ${bestParams.positive2Percent}%")
    print("3+: ${bestParams.positive3Percent}%")
}

// miR-155 stain evaluation function
def evaluateMiR155Stain(detections) {
    def totalCells = detections.size()
    def negativeCells = detections.count { it.getPathClass()?.getName() == 'Negative' }
    def positive1Cells = detections.count { it.getPathClass()?.getName() == '1+' }
    def positive2Cells = detections.count { it.getPathClass()?.getName() == '2+' }
    def positive3Cells = detections.count { it.getPathClass()?.getName() == '3+' }
    
    def negativePercent = (negativeCells / totalCells) * 100
    def positive1Percent = (positive1Cells / totalCells) * 100
    def positive2Percent = (positive2Cells / totalCells) * 100
    def positive3Percent = (positive3Cells / totalCells) * 100
    
    // Calculate stain distribution quality (prefer balanced distribution)
    def stainDistribution = calculateStainDistribution(negativePercent, positive1Percent, positive2Percent, positive3Percent)
    
    return [
        negativePercent: Math.round(negativePercent * 100) / 100,
        positive1Percent: Math.round(positive1Percent * 100) / 100,
        positive2Percent: Math.round(positive2Percent * 100) / 100,
        positive3Percent: Math.round(positive3Percent * 100) / 100,
        stainDistribution: stainDistribution
    ]
}

def calculateStainDistribution(neg, pos1, pos2, pos3) {
    // Prefer distributions where we have some of each category
    // but not too extreme in any direction
    def hasAllCategories = (neg > 5 && pos1 > 5 && pos2 > 5 && pos3 > 5) ? 20 : 0
    def notTooExtreme = (neg < 90 && pos3 < 50) ? 15 : 0
    def reasonablePositive = ((pos1 + pos2 + pos3) > 10 && (pos1 + pos2 + pos3) < 80) ? 15 : 0
    
    return hasAllCategories + notTooExtreme + reasonablePositive
}

// Quality scoring function (customize based on your criteria)
def calculateQualityScore(cellCount, avgNucleusArea, avgOD, stainMetrics) {
    // Example scoring logic - adjust based on your requirements
    def countScore = Math.min(cellCount / 1000.0, 1.0) * 30  // Prefer reasonable cell counts
    def areaScore = (avgNucleusArea > 10 && avgNucleusArea < 100) ? 20 : 0  // Prefer realistic nucleus sizes
    def odScore = (avgOD > 0.05 && avgOD < 2.0) ? 20 : 0  // Prefer meaningful OD values
    def stainScore = stainMetrics.stainDistribution  // Use stain distribution score
    
    return countScore + areaScore + odScore + stainScore
}
    
    return hasAllCategories + notTooExtreme + reasonablePositive
}

// Quality scoring function (customize based on your criteria)
def calculateQualityScore(cellCount, avgNucleusArea, avgOD, stainMetrics) {
    // Example scoring logic - adjust based on your requirements
    def countScore = Math.min(cellCount / 1000.0, 1.0) * 30  // Prefer reasonable cell counts
    def areaScore = (avgNucleusArea > 10 && avgNucleusArea < 100) ? 20 : 0  // Prefer realistic nucleus sizes
    def odScore = (avgOD > 0.05 && avgOD < 2.0) ? 20 : 0  // Prefer meaningful OD values
    def stainScore = stainMetrics.stainDistribution  // Use stain distribution score
    
    return countScore + areaScore + odScore + stainScore
}
