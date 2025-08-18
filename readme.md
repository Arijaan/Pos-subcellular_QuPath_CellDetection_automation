# Info:

This repo contains scripts for automating the process of cell detection and classification for miRNAscope and RNAscope assays. The positive cell detection can also be used for IHC assays. 


Flowchart:
https://unibremende-my.sharepoint.com/:u:/g/personal/arian2_uni-bremen_de/EQiOiY7sl6pBuDFZZ62uWhcB2Img44maO3RIhlykpXyUzQ?e=wMveRc

# Dependencies:

## R-Libraries:

ggplot2, tibble, dplyr, readxl

```R
packages.install("library_name")
```


## Python libraries:

Pandas

```
pip install pandas
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
