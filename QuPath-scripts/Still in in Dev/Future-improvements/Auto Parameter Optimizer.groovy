import qupath.lib.objects.PathAnnotationObject
import qupath.lib.roi.ROIs
import qupath.lib.color.ColorDeconvolutionStains
import qupath.lib.images.servers.ImageServer
import qupath.lib.regions.RegionRequest
import ij.process.AutoThresholder
import javax.swing.JOptionPane
import qupath.lib.images.ImageData

def imageData = getCurrentImageData()
if (imageData == null) {
    print 'No image open.'
    return
}

def server = imageData.getServer()
double px = server.getPixelCalibration().pixelWidthMicrons
if (Double.isNaN(px) || px <= 0)
    px = 0.5 // fallback

// -------- 1. Ask image type & set stains ----------
def choices = ['Brightfield (H-DAB)', 'Brightfield (Other)', 'Keep current'] as Object[]
def choice = JOptionPane.showInputDialog(null, 'Select image type / stain set:', 'Cell detection setup', JOptionPane.PLAIN_MESSAGE, null, choices, choices[0])
if (choice == null) {
    print 'Cancelled.'
    return
}

def setStainsIfNeeded = { name, json ->
    def stains = ColorDeconvolutionStains.parseColorDeconvolutionStainsArg(json)
    imageData.setColorDeconvolutionStains(stains)
    print "Applied stain set: ${name}"
}

switch (choice) {
    case 'Brightfield (H-DAB)':
        setImageType('BRIGHTFIELD_H_DAB')
        // Typical robust H-DAB vectors (adjust if you have calibrated ones)
        setStainsIfNeeded('H-DAB default',
            '{"Name":"H-DAB","Stain 1":"Hematoxylin","Values 1":"0.65 0.70 0.29","Stain 2":"DAB","Values 2":"0.27 0.57 0.78","Background":"255 255 255"}')
        break
    case 'Brightfield (Other)':
        setImageType('BRIGHTFIELD_OTHER')
        // Example: Hematoxylin + FastRed (placeholder second channel named miR-155)
        setStainsIfNeeded('H + miR-155',
            '{"Name":"H-miR","Stain 1":"Hematoxylin","Values 1":"0.77 0.55 0.31","Stain 2":"miR-155","Values 2":"0.39 0.89 0.24","Background":"255 255 255"}')
        break
    default:
        print 'Keeping current image type & stains.'
}

// -------- 2. Reset objects & create full-image annotation ----------
def hier = imageData.getHierarchy()
def removeAll = { objs -> if (objs && !objs.isEmpty()) hier.removeObjects(objs, true) }
removeAll(hier.getDetectionObjects())
removeAll(hier.getAnnotationObjects())
long fullW = server.getWidth()
long fullH = server.getHeight()
def roi = ROIs.createRectangleROI(0, 0, fullW, fullH, null)
def ann = new PathAnnotationObject(roi)
hier.addPathObject(ann)
fireHierarchyUpdate()
print 'Reset annotations & detections; added full-image annotation.'

// -------- 3. Heuristic threshold estimation on Hematoxylin channel ----------
double computeHemaThreshold(ImageData id, ImageServer srv) {
    try {
        int imgWidth  = (int)srv.getWidth()
        int imgHeight = (int)srv.getHeight()
        int downsample = Math.max(1, (int)Math.round(imgWidth / 2000.0))
        int w = imgWidth / downsample
        int h = imgHeight / downsample
        int sampleW = Math.min(w, 1500)
        int sampleH = Math.min(h, 1500)
        int x = (w - sampleW) / 2
        int y = (h - sampleH) / 2
        print "Sampling for threshold: downsample=${downsample}, window=${sampleW}x${sampleH} (orig ${(sampleW*downsample)}x${(sampleH*downsample)})"
    // Use RegionRequest API (BioFormats servers do not expose the direct overload used previously)
    int reqX = x*downsample
    int reqY = y*downsample
    int reqW = sampleW*downsample
    int reqH = sampleH*downsample
    def req = RegionRequest.createInstance(srv.getPath(), downsample, reqX, reqY, reqW, reqH)
    def img = srv.readBufferedImage(req)
        if (img == null) return Double.NaN

        // Build histogram directly (avoid storing all values)
        int[] hist = new int[256]
        double maxVal = 0d
        for (int yy=0; yy<img.getHeight(); yy+=2) {
            for (int xx=0; xx<img.getWidth(); xx+=2) {
                int rgb = img.getRGB(xx, yy)
                int r = (rgb>>16)&0xFF
                int g = (rgb>>8)&0xFF
                int b = rgb&0xFF
                double mean = (r+g+b)/765.0
                double od = -Math.log(Math.max(mean, 1e-4))
                if (od > maxVal) maxVal = od
                // Temporarily store od scaled later; we defer until max known
                // We'll collect in a lightweight list for single pass scaling
            }
        }
        if (maxVal <= 0) return Double.NaN
        // Second pass to populate histogram (need maxVal)
        for (int yy=0; yy<img.getHeight(); yy+=2) {
            for (int xx=0; xx<img.getWidth(); xx+=2) {
                int rgb = img.getRGB(xx, yy)
                int r = (rgb>>16)&0xFF
                int g = (rgb>>8)&0xFF
                int b = rgb&0xFF
                double mean = (r+g+b)/765.0
                double od = -Math.log(Math.max(mean, 1e-4))
                int bin = (int)Math.round(255.0 * (od / maxVal))
                if (bin > 255) bin = 255
                hist[bin]++
            }
        }
        int total = 0; for (int v : hist) total += v
        if (total < 1000) return Double.NaN
        int threshIndex = new AutoThresholder().getThreshold(AutoThresholder.Method.Otsu, hist)
        double fraction = (threshIndex / 255.0)
        double estimated = fraction * maxVal
        return Math.min(0.5, Math.max(0.02, estimated))
    } catch (e) {
        print "Threshold estimation failed: ${e.message}"
        return Double.NaN
    }
}

double threshold = computeHemaThreshold(imageData, server)
if (Double.isNaN(threshold)) {
    threshold = 0.08
    print "Using fallback threshold: ${threshold}"
} else {
    print "Estimated hematoxylin nucleus threshold: ${String.format('%.4f', threshold)}"
}

// -------- 4. Derive size & smoothing parameters ----------
/*
 Revised heuristics:
 - Choose nucleus diameter estimate based on pixel size (px).
 - Use broader allowed range; smaller min area to avoid filtering all cells.
*/
double assumedNucDiam
if      (px < 0.15) assumedNucDiam = 8.0
else if (px < 0.25) assumedNucDiam = 9.0
else if (px < 0.35) assumedNucDiam = 10.0
else                assumedNucDiam = 11.0

double minArea = Math.PI * Math.pow(assumedNucDiam * 0.30, 2)  // less strict
double maxArea = Math.PI * Math.pow(assumedNucDiam * 2.4, 2)

double cellExpansion = Math.min(assumedNucDiam * 0.75, 10.0)
double backgroundRadius = Math.max(assumedNucDiam * 3.0, 20.0)
double medianRadius = 1.0
double sigma = 1.2
double maxBackground = 2.0
print "Heuristics: px=${String.format('%.3f', px)}µm, nucDiam≈${assumedNucDiam}µm, minArea=${String.format('%.1f', minArea)} maxArea=${String.format('%.1f', maxArea)}"

// -------- 5. Adaptive multi-pass Watershed Cell Detection ----------
clearDetections()

def channelName = 'Hematoxylin OD'
def baseParams = [
    'detectionImageBrightfield' : channelName,
    'requestedPixelSizeMicrons' : px,
    'backgroundRadiusMicrons'   : backgroundRadius,
    'medianRadiusMicrons'       : medianRadius,
    'sigmaMicrons'              : sigma,
    'minAreaMicrons'            : minArea,
    'maxAreaMicrons'            : maxArea,
    'threshold'                 : threshold,
    'maxBackground'             : maxBackground,
    'watershedPostProcess'      : true,
    'cellExpansionMicrons'      : cellExpansion,
    'includeNuclei'             : true,
    'smoothBoundaries'          : true,
    'makeMeasurements'          : true
]

def attemptThresholds = [
    baseParams.threshold,
    baseParams.threshold*0.85,
    baseParams.threshold*0.75,
    baseParams.threshold*0.65,
    baseParams.threshold*0.55,
    baseParams.threshold*0.45,
    baseParams.threshold*0.38
]
int bestCount = -1
def bestParams = null
int attempt = 0
double workingMinArea = baseParams.minAreaMicrons
for (t in attemptThresholds) {
    attempt++
    clearDetections()
    def p = new LinkedHashMap(baseParams)
    p.threshold = t
    // gradually relax min area on later attempts
    if (attempt > 1) {
        workingMinArea *= 0.85
        p.minAreaMicrons = workingMinArea
    }
    print "Attempt ${attempt}: threshold=${String.format('%.4f', p.threshold)} minArea=${String.format('%.1f', p.minAreaMicrons)}"
    runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', p)
    int count = getCellObjects().size()
    print " -> Detected ${count} cells"
    if (count > bestCount) {
        bestCount = count
        bestParams = p
    }
    // Early stop if raising detection didn't help much after first two attempts
    if (attempt > 2 && count < bestCount * 1.05 && bestCount > 0) {
        print 'Early stopping: further threshold lowering gives limited gains.'
        break
    }
}

// If last run wasn't best, rerun with best params to retain those detections
if (bestParams != null) {
    clearDetections()
    print 'Re-running with best parameters:'
    bestParams.each { k,v ->
        String valStr
        if (v instanceof Number) {
            try { valStr = String.format('%.3f', (v as Number).doubleValue()) } catch (Exception ignore) { valStr = v.toString() }
        } else valStr = v.toString()
        println "  ${k} = ${valStr}"
    }
    runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', bestParams)
    print "Final cell count: ${getCellObjects().size()}"
} else {
    print 'No detection attempts succeeded.'
}