# Version:

0.0.1

# Info:

This repo contains scripts for automating the process of cell detection and classification for miRNAscope and RNAscope assays. The positive cell detection can also be used for IHC assays. 

The QuPath plugins currently only contain finished folders for counting and analyzing the measurements of pancreatic and spleen cells specifically.

As of the current version, scripts need to be ran individually and staining vectors set manually. All variable names should also be set according to your project. This will be changed in later versions to implement more simple GUI prompts.

 The workflow starts in groovy with manual selection of representative staining vectors and cell detection parameters and then adjusting the appropriate project automation script and running it. Optimization for single images might be needed in this step. Afterwards, in the save measurements script in the project automation folder can be run to save the number of cells with classifications in comma-separated .csv files (Change the path to a directory of your choice). Afterwards the python script "measurement summarizer" can be ran to get summary files including the sum of all cells in your project (don't forget appropriate naming). The R-files can then be ran to get plots based on the measurement summary files, outputting barplots and prop test results.


Flowchart for optimizing cell detection:

https://unibremende-my.sharepoint.com/:u:/g/personal/arian2_uni-bremen_de/EQiOiY7sl6pBuDFZZ62uWhcB2Img44maO3RIhlykpXyUzQ?e=wMveRc

# Dependencies:

## R-Libraries:

ggplot2, tibble, dplyr, readxl

```R
packages.install("library_name")
```


## Python libraries:

Python 3.0 and above

Pandas
tkinter

```
pip install pandas
```
```
pip install tkinter
```




# Abbreviations/Naming System:

PA: Project automation

F: Final

T: in testing phase

Col: Color convolution included

AI: Completely Written with generative AI and not tested

# Future updates:

⦁	Automatic selection of cell detection parameters for each image

⦁	Cleaning up the workflow by using main-scripts, out of which functions are read out

⦁	Improving the save measurements option, expanding on customizability

⦁	Improving the flow of performing each analysis with integrating scripts into each other
