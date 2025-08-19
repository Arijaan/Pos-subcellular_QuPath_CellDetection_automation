// For adding plugins from single files and applying it to multiple images/files in a project

// instead of defining the function, the script,
//  containing the desired plugins/parameters are loaded into the loop. NOT TESTED/IMPLEMENTED

def scriptPath = "path/to/your/script.groovy"
runScript(new File(scriptPath))

// Or use relative paths from the script directory
runScript(new File(getQuPath().getScriptDirectory(), "1-subcellular-spot-detection.groovy"))