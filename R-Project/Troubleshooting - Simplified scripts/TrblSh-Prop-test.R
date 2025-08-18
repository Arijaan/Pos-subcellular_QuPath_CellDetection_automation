library(ggplot2)
library(dplyr)
library(tidyr)
library(tibble)

# Enter Data

path1 <- "C:/Users/psoor/OneDrive/Desktop/Bachelorarbeit/Measurements/Control - 2.csv"
path2 <- "C:/Users/psoor/OneDrive/Desktop/Bachelorarbeit/Measurements/T1D - 2.csv"


# Name Data

c1 <- 5
c2 <- 6
control <- read.csv(path1)[c(1:7), c(c1, c2)]
t1d <- read.csv(path2)[c(1:7), c(c1, c2)]

# Debug Name Data 
control[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(control[, 2])))
t1d[, 2] <- as.numeric(gsub("[^0-9.-]", "", as.character(t1d[, 2])))
control <- control[complete.cases(control), ]
t1d <- t1d[complete.cases(t1d), ]


combined_data <- rbind(control, t1d)


all_classifications <- unique(combined_data$classification)

t1d_total <- read.csv(path1)[2, c2]
t2d_total <- read.csv(path2)[2, c2]

# define T1D and ctrl pos counts
