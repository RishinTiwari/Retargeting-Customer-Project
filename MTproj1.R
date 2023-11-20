#reloading the existing libraries
library("writexl")
library("stringr")
library("dplyr")
library("readr")
library("rio")
library("moments")
library(stargazer)
#install.packages("writexl")
library("writexl")
library("openxlsx")

#loading csv file
abddata <- read.csv("/Users/rishintiwari/Desktop/Fall 2023/Analytical Methods of Business(AMB)/Midterm Project/Abandoned.csv" , header=T, na.strings="")
resdata <- read.csv("/Users/rishintiwari/Desktop/Fall 2023/Analytical Methods of Business(AMB)/Midterm Project/Reservation.csv" , header=T, na.strings="")

# displaying the column names and dimensions 
variable_names <- names(abddata)
print(variable_names)
dim(abddata)
dim(resdata)


# Matching data
match_email <- abddata$Email[complete.cases(abddata$Email)] %in% resdata$Email[complete.cases(resdata$Email)]
match_incoming <- abddata$Incoming_Phone[complete.cases(abddata$Incoming_Phone)] %in% resdata$Incoming_Phone[complete.cases(resdata$Incoming_Phone)]
match_contact <- abddata$Contact_Phone[complete.cases(abddata$Contact_Phone)] %in% resdata$Contact_Phone[complete.cases(resdata$Contact_Phone)]
match_incoming_contact <- abddata$Incoming_Phone[complete.cases(abddata$Incoming_Phone)] %in% resdata$Contact_Phone[complete.cases(resdata$Contact_Phone)]
match_contact_incoming <- abddata$Contact_Phone[complete.cases(abddata$Contact_Phone)] %in% resdata$Incoming_Phone[complete.cases(resdata$Incoming_Phone)]

# Create flags 
# Create match_email and email columns
abddata$match_email <- 0 
abddata$match_email[complete.cases(abddata$Email)] <- 1* match_email

# Create match_incoming and incoming columns
abddata$match_incoming <- 0 
abddata$match_incoming[complete.cases(abddata$Incoming_Phone)] <- 1* match_incoming

# Create match_contact and contact columns
abddata$match_contact <- 0 
abddata$match_contact[complete.cases(abddata$Contact_Phone)] <- 1* match_contact

# Create match_incoming_contact column
abddata$match_incoming_contact  <- 0
abddata$match_incoming_contact[complete.cases(abddata$Incoming_Phone) ] <- 1* match_incoming_contact

# Create match_contact_incoming column
abddata$match_contact_incoming  <- 0
abddata$match_contact_incoming[complete.cases(abddata$Contact_Phone) ] <- 1* match_contact_incoming

abddata$pur <- 0
abddata$pur <- 1 * (abddata$match_email | abddata$match_incoming | abddata$match_contact | abddata$match_incoming_contact | abddata$match_contact_incoming)

#Marking Test as 1 and Control as 0 for Treatment variables 
abddata[abddata ==  'test'] <- 1
abddata[abddata ==  'control'] <- 0

#Cross Tabluation Value for the whole dataset
#test & purchase
a1 = nrow(abddata[abddata$pur == '1' & abddata$Test_Control == '1', ]) 
#test & not purchase
b1 = nrow(abddata[abddata$pur == '0'  & abddata$Test_Control == '1', ]) 
# control & purchase
c1 = nrow(abddata[abddata$pur == '1'  & abddata$Test_Control == '0', ]) 
#control & not purchase
d1 = nrow(abddata[abddata$pur == '0'  & abddata$Test_Control == '0', ])

data.matrix = matrix(c(a1,b1,c1,d1),ncol=2,nrow=2,byrow=TRUE)
colnames(data.matrix) = c("Purchased","Not Purchased")
rownames(data.matrix) = c("Treatment","Control")
data.matrix.tb <- as.table(data.matrix)
data.matrix.tb

# Create a function to compute cross-tabulations for a given state
funct_cross_tab <- function(data, state) {
  subset_state <- subset(data, Address == state)
  cross_tab <- table(subset_state$Test_Control, subset_state$pur)
  return(cross_tab)
}

# Example usage for different states
cross_tab_CA <- funct_cross_tab(abddata, "CA")
print(cross_tab_CA, quote = FALSE)
cross_tab_LA <- funct_cross_tab(abddata, "LA")
print(cross_tab_LA, quote = FALSE)
cross_tab_FL <- funct_cross_tab(abddata, "FL")
cross_tab_OH <- funct_cross_tab(abddata, "OH")
cross_tab_UT <- funct_cross_tab(abddata, "UT")

# creating Dataframe with specific column for Exporting to our PC
final.data = select(abddata, c('Caller_ID','Test_Control','pur', 'Address' , 'Email') , )
Final.data.sub.1 = subset(final.data, pur == 1 )
Final.data.sub.0 = subset(final.data, pur == 0)
# Merge the two data frames into one
combined_data <- rbind(Final.data.sub.1, Final.data.sub.0)
# Save the combined data to an Excel file
write_xlsx(combined_data, "/Users/rishintiwari/Desktop/final_data_sheet_combined.xlsx")


# Marking Email and address as 1 in case of values available 
abddata$Email<- 1*complete.cases(abddata$Email)
abddata$Address <- 1*complete.cases(abddata$Address)

#Analyzing Test-Control Division
abddata$Test_Control <- as.numeric(abddata$Test_Control)
summary(abddata$Test_Control)
table(abddata$Test_Control)


#Linear regression for Pur based on Test_Control
out_lm1 <- lm(pur ~ Test_Control, data = abddata)
summary(out_lm1 )
#The low R-squared value (0.01745) indicates that the 'Test_Control' variable explains only a small proportion of the variance in 'pur,' suggesting that there may be other factors not accounted for in this model.
#Intercept (Intercept): 0.022270
#Coefficient for Test_Control1 (Test_Control1): 0.058602

#Linear regression for Pur based on Test_control + Email + Address 
out_lm2 <- lm(pur ~ Test_Control + Email + Address , data = abddata)
summary(out_lm2)
#The model is statistically significant and explains a small portion of the variance in 'pur' (2.268%).
#Intercept (Intercept): 0.010703
#Coefficient for Test_Control1 (Test_Control1): 0.057623
#Coefficient for Email (Email): 0.036416
#Coefficient for Address (Address): 0.016873

#Linear regression for pur  based on Test_control * Email + Test_control * Address  
out_lm3 <- lm(pur ~ Test_Control* Email + Test_Control* Address , data = abddata)
summary(out_lm3)
#Intercept (Intercept): 0.016844
#Coefficient for Test_Control1 (Test_Control1): 0.045338
#Coefficient for Email (Email): 0.007599
#Coefficient for Address (Address): 0.010301
#Coefficient for the interaction between Test_Control1 and Email (Test_Control1:Email): 0.052981
#Coefficient for the interaction between Test_Control1 and Address (Test_Control1:Address): 0.013011
#The R-squared value of 0.02466 indicates that the model explains only a small portion of the variance in the "pur" variable, suggesting that other factors not included in the model may influence purchase decisions.

#ANOVA 
abddata$Test_Control  = as.factor(abddata$Test_Control)
abddata$pur  = as.numeric(abddata$pur)

model <- aov( pur~Test_Control, data = abddata)
summary(model)
#The ANOVA results demonstrate that the 'Test_Control' variable significantly affects the 'pur' variable. 
#The small p-value (< 0.05) provides strong evidence to reject the null hypothesis, indicating that 'Test_Control' has a substantial impact on 'pur'. 
#The high F-statistic value (149.9) further supports the significance of this relationship.
# Generate summary table 
stargazer(out_lm1, out_lm2, out_lm3, type = "html", out = "midterm_proj.html")

