#Load Libraries Needed
library(xlsx)#package to read directly from Excel
library(tidyverse)#data frame operations package
library(ggplot2)#library for plotting
library(lattice)#library for plotting grouped bar charts
library(lessR)#library for easy bar charts
library(grid)#to plot multiple graphs
library(gridExtra)#to plot multiple graphs
library(Hmisc)#for data frame statistics
library(reshape2)#to modify data for correlation analysis
library(caret)#for confusion matrix
library(InformationValue)#for confusion matrix
library(rpart.plot)#for decision tree
library(rpart)#for decision tree

#Read in file
#command comes from the xlsx package, filename and then the ,1 means the first sheet
df<-read.xlsx("C:/Users/MarkCichonski/Documents/Data Science Projects/Data Science Greenbelt/20221205 Intro Project/1. Original Data/marketing_campaign.xlsx",1)

#Data EDA and Preprocessing (EDA is Exploratory Data Analysis)
dim(df)#dimensions of the df dataframe

glimpse(df)#shows the columns with some sample data

apply(is.na(df),2, which)#shows missing values
#Here we have 3 categorical variables and 26 numerical variables
#Only Income variable has missing values

# check if the target is imbalanced
#do a barchart, standard R functionality of the values in the response column
barplot(table(df$Response),xlab="Response",ylab="Count",main="Balance of Responses")

# fetch age from Year_Birth
#calculate age from today's date and add a new column
df$age=(as.integer(format(Sys.Date(),"%Y")))-df$Year_Birth
#drop the year of birth column
df$Year_Birth<-NULL

#Income Spending by Age #Scatterplot
#Simple Scatterplot using R functionalty (vs a second graphing package)
plot(df$age, df$Income, main="Scatterplot Example",
     xlab="Age ", ylab="Income ")
#simple boxplot
boxplot(df$age)
#nicer scatterplot using ggplot
ggplot(df, aes(x=age, y=Income, color=Response)) + 
  geom_point(size=1) 

#Seems there are few outliers in Age as well as Income
#Income for most of the population is under 100k
#Most of the people in the dataset aged between 45-65

# calculate customers spending
df$spending<-df$MntFishProducts+df$MntFruits+df$MntGoldProds+df$MntSweetProducts+df$MntWines+df$MntMeatProducts
#drop the individual spending columns
df$MntFishProducts<-NULL
df$MntFruits<-NULL
df$MntGoldProds<-NULL
df$MntMeatProducts<-NULL
df$MntSweetProducts<-NULL
df$MntWines<-NULL

#plot income vs. spending
ggplot(df, aes(x=Income, y=spending, color=Response)) + 
  geom_point(size=1) 
hist(df$spending)

#Most of the people who responded have high spending
#The number of people whose spending is less than 200 are high in numbersMost of the people who responded have high spending

# Income and spending by education
#barplot education vs. income
p<-ggplot(data=df, aes(x=Education, y=Income)) +
  geom_bar(stat="identity")
p

#boxplot education vs. income
p1<-ggplot(data=df, aes(x=Education, y=Income)) +
  geom_boxplot()
p1

#barplot education vs. spending
p2<-ggplot(data=df, aes(x=Education, y=spending)) +
  geom_bar(stat="identity")
p2
#boxplot education vs. spending
p3<-ggplot(data=df, aes(x=Education, y=spending)) +
  geom_boxplot()
p3

#Graduate spending is higher according to their income

#Let's encode the education variable for modelling later
encode_ordinal <- function(x, order = unique(x)) {
  x <- as.numeric(factor(x, levels = order, exclude = NULL))
  x
}
#can you make a library of functions available to all scripts?
#Create a new data frame to check function
new_df<-df
new_df[["education_encoded"]] <- encode_ordinal(df[["Education"]])
#drop test data frame
rm(new_df)#rm is the command for remove

#Reapply function on original data frame
df[["education_encoded"]] <- encode_ordinal(df[["Education"]])
#drop Education column
df$Education<-NULL

# Income and spending by marital status
#barplot income by marital status
#Use trellis barcode
barchart(Income ~ Marital_Status, data=df, groups=Response)
#boxplot income by marital status
ggplot(data=df, aes(x=Marital_Status, y=Income, fill=factor(Response))) +
  geom_boxplot()
#barplot spending by marital status
barchart(spending ~ Marital_Status, data=df, groups=Response)
#boxplot spending by marital status
ggplot(data=df, aes(x=Marital_Status, y=spending, fill=factor(Response))) +
  geom_boxplot()

#let's manually encode the marital status
df$Marital_Status[df$Marital_Status=="Married"]<-0
df$Marital_Status[df$Marital_Status=="Together"]<-0
df$Marital_Status[df$Marital_Status=="Single"]<-1
df$Marital_Status[df$Marital_Status=="Divorced"]<-1
df$Marital_Status[df$Marital_Status=="Widow"]<-1
df$Marital_Status[df$Marital_Status=="Alone"]<-1
df$Marital_Status[df$Marital_Status=="Absurd"]<-1
df$Marital_Status[df$Marital_Status=="YOLO"]<-1

#Income and spending by kids
#barplot income by kids
#Use lessR BarChart
BarChart(data=df, Kidhome, Income, by=Response, beside=TRUE, values_position = "out", stat = "mean")
#boxplot income by kids
ggplot(df,aes(factor(Kidhome),Income))+geom_boxplot(aes(fill=(factor(Response))))
#barplot spending by kids
BarChart(data=df, Kidhome, spending, by=Response, beside=TRUE, values_position = "out", stat = "mean")
#boxplot spending by marital status
ggplot(df,aes(factor(Kidhome),spending))+geom_boxplot(aes(fill=(factor(Response))))

#barplot income by teens
#Use lessR BarChart
BarChart(data=df, Teenhome, Income, by=Response, beside=TRUE, values_position = "out", stat = "mean")
#boxplot income by kids
ggplot(df,aes(factor(Teenhome),Income))+geom_boxplot(aes(fill=(factor(Response))))
#barplot spending by kids
BarChart(data=df, Teenhome, spending, by=Response, beside=TRUE, values_position = "out", stat = "mean")
#boxplot spending by marital status
ggplot(df,aes(factor(Teenhome),spending))+geom_boxplot(aes(fill=(factor(Response))))

#lets change the factor for kids and teens to yes or no (0 or 1)
df$has_kid[df$Kidhome=="0"]<-0
df$has_kid[df$Kidhome=="1"]<-1
df$has_kid[df$Kidhome=="2"]<-1
df$has_teen[df$Teenhome=="0"]<-0
df$has_teen[df$Teenhome=="1"]<-1
df$has_teen[df$Teenhome=="2"]<-1
#drop Kidhome and Teenhome
df$Kidhome<-NULL
df$Teenhome<-NULL

#now let's look at the by number of purchases
g1=ggplot(data=df, aes(x=NumDealsPurchases, y=Income))+geom_bar(stat="identity")
g2=ggplot(data=df, aes(x=NumDealsPurchases, y=spending))+geom_bar(stat="identity")
g3=ggplot(data=df, aes(x=NumWebPurchases, y=Income))+geom_bar(stat="identity")
g4=ggplot(data=df, aes(x=NumWebPurchases, y=spending))+geom_bar(stat="identity")
g5=ggplot(data=df, aes(x=NumCatalogPurchases, y=Income))+geom_bar(stat="identity")
g6=ggplot(data=df, aes(x=NumCatalogPurchases, y=spending))+geom_bar(stat="identity")
g7=ggplot(data=df, aes(x=NumStorePurchases, y=Income))+geom_bar(stat="identity")
g8=ggplot(data=df, aes(x=NumStorePurchases, y=spending))+geom_bar(stat="identity")
g9=ggplot(data=df, aes(x=NumWebVisitsMonth, y=Income))+geom_bar(stat="identity")
g10=ggplot(data=df, aes(x=NumWebVisitsMonth, y=spending))+geom_bar(stat="identity")
grid.arrange(g1,g2,g3,g4,g5,g6,g7,g8,g9,g10, ncol=2, top="Income and Spending by Purchase Type")

#In most of the cases Income has no effect on the number of purchases however the spending increases as the number of purchases increases
#We can also conclude that the people who visits store more frequently spends higher regardless the income
#Also we can conclude that the frequent visitors has low income as well as low spending, it means the new user's spending are high

#Let's convert customer date to date time
#convert the date of enrollment to datetime
df$Dt_Customer<-as.POSIXct(df$Dt_Customer)

#creating features from date of enrollment
df$Dt_Customer<-as.Date(df$Dt_Customer,"%Y-%m-%d")
df$Year_Customer<-substring(df$Dt_Customer,1,4)
df$Month_Customer<-substring(df$Dt_Customer,6,7)
df$Day_Customer<-substring(df$Dt_Customer,9,10)
df$Dt_Customer<-NULL

#Plot income and spending by time factors
h1=ggplot(data=df, aes(x=Year_Customer, y=Income))+geom_bar(stat="identity")
h2=ggplot(data=df, aes(x=Year_Customer, y=spending))+geom_bar(stat="identity")
h3=ggplot(data=df, aes(x=Month_Customer, y=Income))+geom_bar(stat="identity")
h4=ggplot(data=df, aes(x=Month_Customer, y=spending))+geom_bar(stat="identity")
h5=ggplot(data=df, aes(x=Day_Customer, y=Income))+geom_bar(stat="identity")
h6=ggplot(data=df, aes(x=Day_Customer, y=spending))+geom_bar(stat="identity")
grid.arrange(h1,h2,h3,h4,h5,h6, ncol=2, top="Income and Spending by Enrollment")

#let's quickly look at the data frame
desc<-describe(df)
desc
#min and max value is same for few variables, lets check it has only one value
#check outliers in Age and Income

#Check which columns have the same data for all rows
i1 <- sapply(df, function(x) length(unique(x)) ==1)
df[i1]
#Let's remove these two variables
df$Z_CostContact<-NULL
df$Z_Revenue<-NULL

#Outliers Handling
#check outliers in age variable
i1=ggplot(df, aes(x=age)) + 
  geom_histogram(aes(y=..density..),      # Histogram with density instead of count on y-axis
                 binwidth=.5,
                 colour="black", fill="white") +
  geom_density(alpha=.2, fill="#FF6666")
i2=ggplot(df,aes(age))+geom_boxplot()
grid.arrange(i1,i2, ncol=2, top="Outliers in the Age Variable")

# lets check the Income variable for outliers
j1=ggplot(df, aes(x=Income)) + 
  geom_histogram(aes(y=..density..),      # Histogram with density instead of count on y-axis
                 binwidth=.5,
                 colour="black", fill="white") +
  geom_density(alpha=.2, fill="#FF6666")
j2=ggplot(df,aes(Income))+geom_boxplot()
grid.arrange(j1,j2, ncol=2, top="Outliers in the Income Variable")

# remove the outliers before handling the missing values
df<-subset(df,df$Income<200000)
df<-subset(df, df$age<100)

#fill income null values with median
df$Income[is.na(df$Income)] <- median(df$Income, na.rm=TRUE)

#let's check for duplicate rows
sum(duplicated(df))

# lets check the correlation between variables through heatmap
#let's make sure all the variables are numeric at this time
df<-as.data.frame(apply(df,2,as.numeric))#this converts frame to all numeric
sapply(df, class)#this checks that conversion was successful
#Correlation matrix can be created using the R function cor() :
cormat <- round(cor(df),2)
head(cormat)
#The package reshape is required to melt the correlation matrix :
melted_cormat <- melt(cormat)
head(melted_cormat)
#The function geom_tile()[ggplot2 package] is used to visualize the correlation matrix :
ggplot(data = melted_cormat, aes(x=Var1, y=Var2, fill=value)) + 
  geom_tile()
#fancier heatmap
ggplot(data = melted_cormat, aes(Var2, Var1, fill = value))+
  geom_tile(color = "white")+
  scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
                       midpoint = 0, limit = c(-1,1), space = "Lab", 
                       name="Pearson\nCorrelation") +
  theme_minimal()+ 
  theme(axis.text.x = element_text(angle = 45, vjust = 1, 
                                   size = 12, hjust = 1))+
  coord_fixed()

#quick check of correlation coeff.
melted_cormat[order(-melted_cormat$value),]#this prints out the correlation coefficient
#notice the "-" sign.  That means the list sorts from highest to lowest

#make this example reproducible
set.seed(1)

#this splits the data frame 70/30 for training and testing which is a rule of thumb
sample <- sample(c(TRUE, FALSE), nrow(df), replace=TRUE, prob=c(0.7,0.3))
#what does replace mean in this context?
train  <- df[sample, ]
test   <- df[!sample, ]


#Modeling
#Logistic Regression
model <- glm(Response ~.,family=binomial(link='logit'),data=train)
summary(model)
fitted.results <- predict(model,newdata=subset(test,),type='response')
fitted.results <- ifelse(fitted.results > 0.5,1,0)
misClasificError <- mean(fitted.results != test$Response)
print(paste('Accuracy',1-misClasificError))

#confusion matrix
#find optimal cutoff probability to use to maximize accuracy
optimal <- optimalCutoff(test$Response, fitted.results)[1]
#create confusion matrix
confusion_log<-confusionMatrix(test$Response, fitted.results)
print(confusion_log,mode="everything", printStats="TRUE")

#analyze confusion matrix
#calculate sensitivity
sensitivity(test$Response, fitted.results)
#calculate specificity
specificity(test$Response, fitted.results)
#calculate total misclassification error rate
misClassError(test$Response, fitted.results, threshold=optimal)

#decision tree
#create decision tree
tree<-rpart(Response~.,data=train)
rpart.plot(tree)
printcp(tree)
#confusion matrix on predictions
ptree<-predict(tree,test)
confusion_tree<-confusionMatrix(test$Response, ptree)
confusion_tree

#analyze confusion matrix
#calculate sensitivity
sensitivity(test$Response, tree$y)
#calculate specificity
specificity(test$Response, tree$y)
#calculate total misclassification error rate
misClassError(test$Response, tree$y)
