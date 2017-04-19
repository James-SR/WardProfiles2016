library(ggplot2)
library(reshape2)
library(xlsx)
library(dplyr)
library(gridExtra)
library(grid)

#reads the ward data in as a, then the CE average figures as seperate object ce
a <- read.xlsx('C:/Summary File - Includes Percentiles.xlsx', "Summary Sheet ", startRow = 2, endRow = 54)
ce <- read.xlsx('C:/Summary File - Includes Percentiles.xlsx', "Summary Sheet ", startRow = 54, endRow = 55)

#Function to select only the columns we are interested in give them a friendly name
sel_and_rename <- function(x) {
  select(x, 1, "NEET" = 13, 
         "Average Household Income" = 14, 
         "Claimant Count" = 17, 
         "Key Stage 4 inc Eng. & Maths" = 19, 
         "Key Stage 5 Avg. Point Score" = 20, 
         "Total Crime Rate" = 22, 
         "Burglary - Proportion of total" = 24, 
         "Robbery - Proportion of total" = 26, 
         "Vehicle - Proportion of total" = 28, 
         "Violence - Proportion of total" = 30, 
         "Crimal Damage - Proportion of total" = 32, 
         "Other - Proportion of total" = 34, 
         "Second Homes" = 53, 
         "Single Occupier Properties"= 55, 
         "Long Term Empty or Unfurnished Properties" = 57, 
         "Student Homes" = 59, 
         "Social Rented Properties" = 62, 
         "Fuel Poverty" = 64)
}
#Rename our two datasets
ce <- sel_and_rename(ce)
b <- sel_and_rename(a)
colnames(ce)[1] <- "CheshireEast"

b_scaled <- scale(b[, 2:19], center = TRUE, scale = TRUE)
b_scaled <- as.data.frame(b_scaled)
b_scaled$Ward <- b[,1] #append ward name to numeric values

bmelt <- melt(b_scaled, measure.vars = 1:18)
bmelt$variable <- as.character(bmelt$variable)
bmelt <- tbl_df(bmelt)

bmelt <- bmelt %>%
  mutate(Theme = ifelse(variable == "NEET" | variable == "Average Household Income" | variable == "Claimant Count" | variable == "Fuel Poverty", "Economic",
                        ifelse(variable == "Social Rented Properties" | variable == "Second Homes" | variable == "Student Homes" | variable == "Single Occupier Properties" | variable == "Long Term Empty or Unfurnished Properties", "Housing",
                               ifelse(variable == "Key Stage 5 Avg. Point Score" | variable == "Key Stage 4 inc Eng. & Maths", "Education", "Crime")))) %>%
  arrange(desc(value))
#Order the results as per the table produced later#
bmelt$variable <- as.factor(bmelt$variable)
levels(bmelt$variable) <- c("Fuel Poverty",
                          "Social Rented Properties",
                          "Student Homes",
                          "Long Term Empty or Unfurnished Properties",
                          "Single Occupier Properties",
                          "Second Homes",
                          "Other - Proportion of total",
                          "Crimal Damage - Proportion of total",
                          "Violence - Proportion of total",
                          "Vehicle - Proportion of total",
                          "Robbery - Proportion of total",
                          "Burglary - Proportion of total" ,
                          "Total Crime Rate",
                          "Key Stage 5 Avg. Point Score",
                          "Key Stage 4 inc Eng. & Maths",
                          "Claimant Count",
                          "Average Household Income",
                          "NEET"
)

bmelt <- bmelt[, c(4, 1:3)]

#filter results just to show the ward in question
ward_name <- "Broken Cross and Upton"
w <- filter(bmelt, Ward == ward_name)

m <- ggplot(data = bmelt, aes(variable, value, fill = Theme)) +
  geom_boxplot(outlier.shape = 19) +
  geom_point(data = w, aes(variable, value, shape = 17), size = 3) +
  scale_shape_identity() +
  coord_flip() +
  ylab('Standardised Range') +
  xlab('') +
  theme(axis.text.x = element_blank(), panel.grid.minor=element_blank(), legend.position ="bottom", axis.text=element_text(size=12), legend.text=element_text(size=12)) +
  scale_fill_manual(values = c("#AA272F", "#0075B0", "#FFA100", "#92D400"),
                    breaks = c("Crime", "Economic", "Education", "Housing")) +
  ggtitle(ward_name)                        
print(m)

#--Table creation--#

#create a function to calculate the extremes and apply to the data
extremes <- function(x) {
  c(min = min(x), max = max(x))
}
t <- sapply(b[-1], extremes)
t <- as.data.frame(t)
#filter by ward then add in ward data and ce data, naming rows
w1 <- dplyr::filter(b, Ward.Name == ward_name)
sum_tbl <- rbind(t, w1[, 2:19])
sum_tbl <- rbind(sum_tbl, ce[, 2:19])
rownames(sum_tbl)[1] <- "CE Lowest"
rownames(sum_tbl)[2] <- "CE Highest"
rownames(sum_tbl)[3] <- ward_name
rownames(sum_tbl)[4] <- "Cheshire East"

#Now we format the values for final presentation
sum_tbl[,1] <- paste(format(round(sum_tbl[,1]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,3] <- paste(format(round(sum_tbl[,3]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,4] <- paste(format(round(sum_tbl[,4]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,13] <- paste(format(round(sum_tbl[,13]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,14] <- paste(format(round(sum_tbl[,14]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,15] <- paste(format(round(sum_tbl[,15]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,16] <- paste(format(round(sum_tbl[,16]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,18] <- paste(format(round(sum_tbl[,18]*100, 1), nsmall = 1),"%",sep="")
sum_tbl[,7] <- paste(format(round(sum_tbl[,7], 1), nsmall = 1),"%",sep="")
sum_tbl[,8] <- paste(format(round(sum_tbl[,8], 1), nsmall = 1),"%",sep="")
sum_tbl[,9] <- paste(format(round(sum_tbl[,9], 1), nsmall = 1),"%",sep="")
sum_tbl[,10] <- paste(format(round(sum_tbl[,10], 1), nsmall = 1),"%",sep="")
sum_tbl[,11] <- paste(format(round(sum_tbl[,11], 1), nsmall = 1),"%",sep="")
sum_tbl[,12] <- paste(format(round(sum_tbl[,12], 1), nsmall = 1),"%",sep="")
sum_tbl[,17] <- paste(format(round(sum_tbl[,17], 1), nsmall = 1),"%",sep="")
sum_tbl[,5] <- format(round(sum_tbl[,5], 1), nsmall = 1)
sum_tbl[,6] <- format(round(sum_tbl[,6], 1), nsmall = 1)
sum_tbl[,2] <- paste("£",format(sum_tbl[,2], big.mark=","),sep="")


#The following section transposes the table
summary <- t(sum_tbl)
summary <- as.data.frame(summary)
grid.table(summary)

#Finally we add the transposed table next to the plot

tt <- ttheme_default(colhead=list(fg_params = list(parse=TRUE)))
tbl <- tableGrob(summary, theme=tt)
grid.arrange(m, tbl, ncol=2)

#---scratch area---#