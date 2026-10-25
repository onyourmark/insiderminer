# InsiderMiner, December 2014.
# R translation of the Excel macro step that aligns tweet times with the
# New York trading day and folds weekend tweet activity into Sunday.
# Transcribed from the Google Doc "code for macro done in R".

S.longtermtest = "GRPN.csv"

ToyotaTweets <- read.csv(S.longtermtest)

# 1.2. remove unnecessary columns from ToyotaTweets
# reason: we just need the text and the date for the objective
ToyotaTweets$Tweet.ID <- NULL
ToyotaTweets$Username <- NULL
ToyotaTweets$Location <- NULL
ToyotaTweets$Followings <- NULL
ToyotaTweets$Followers <- NULL

# create one more column to save the new Date (after changing the time zone and subtracting 8 hours)
ToyotaTweets["editDate"] <- NA
ToyotaTweets <- ToyotaTweets[, c(1, 3, 2)]
# change the columns' position, so that the 2 columns with the old date and new date will be next to each other

# 1.3 change the time zone and minus 8 hours of ToyotaTweets
# why? get the tweets' time zone and the stock's time zone to be synchronized
# why? using the tweets that posted earlier than 8:00 AM

# read the Tweets date & time as "GMT" time zone and put that in temp
temp <- as.POSIXct(ToyotaTweets$Date, tz = "GMT", format = "%m/%d/%Y %H:%M")

# temp - 28800 seconds (8 hours), and change that result to America/New_York time zone
# this time is now synchronized with the stock time
ToyotaTweets$editDate <- format(temp - 28800, tz = "America/New_York", usetz = FALSE, format = "%m/%d/%Y")

x = table(ToyotaTweets$editDate)
xdates = names(x)
xcounts = as.vector(x)
toyoDateCount.df = data.frame(xdates, xcounts)
toyoDateCount.df = toyoDateCount.df[order(as.Date(toyoDateCount.df$xdates, format = "%m/%d/%Y")), ]
head(toyoDateCount.df)

OFResult = toyoDateCount.df

# 4.1 - Preparation
# OFResult$checked prevents a weekend from being averaged twice
OFResult$checked <- FALSE
colnames(OFResult)[1] = "date"

# 4.3 - Put the mean of the weekend into Sunday
# (1) if Fri - Sat - Sun then Sun <- mean(Fri, Sat, Sun)  // replace value of Sunday
# (2) if Fri - Sat       then Sun <- mean(Fri, Sat)       // rbind a new Sunday
# (3) if Fri - Sun       then Sun <- mean(Fri, Sun)       // replace value of Sunday
# (4) if Fri only        then Sun <- Fri                  // rbind a new Sunday
# (5) if Sat - Sun       then Sun <- mean(Sat, Sun)       // replace value of Sunday
# (6) if Sat only        then Sun <- Sat                  // rbind a new Sunday
# (7) if Sun only        do nothing

# first loop runs cases (1), (3), (5), (7)
for (i in 1:nrow(OFResult)) {
  if ((weekdays(as.Date(OFResult$date[i], "%m/%d/%Y")) == "Friday") & (OFResult$checked[i] == FALSE)) {
    if ((weekdays(as.Date(OFResult$date[i + 1], "%m/%d/%Y")) == "Saturday") & (OFResult$checked[i + 1] == FALSE)
        & (weekdays(as.Date(OFResult$date[i + 2], "%m/%d/%Y")) == "Sunday") & (OFResult$checked[i + 2] == FALSE)) {
      # case (1)
      j <- i + 2
      OFResult[j, 2] <- c(round(mean(OFResult[i:j, 2])))
      OFResult[i:j, 3] <- TRUE
    }
    else if ((weekdays(as.Date(OFResult$date[i + 1], "%m/%d/%Y")) == "Sunday") & (OFResult$checked[i + 1] == FALSE)) {
      # case (3)
      j <- i + 1
      OFResult[j, 2] <- c(round(mean(OFResult[i:j, 2])))
      OFResult[i:j, 3] <- TRUE
    }
  }
  # case (5)
  if ((weekdays(as.Date(OFResult$date[i], "%m/%d/%Y")) == "Saturday") & (OFResult$checked[i] == FALSE)) {
    if ((weekdays(as.Date(OFResult$date[i + 1], "%m/%d/%Y")) == "Sunday") & (OFResult$checked[i + 1] == FALSE)) {
      j <- i + 1
      OFResult[j, 2] <- c(round(mean(OFResult[i:j, 2])))
      OFResult[i:j, 3] <- TRUE
    }
  }
  # case (7)
  if ((weekdays(as.Date(OFResult$date[i], "%m/%d/%Y")) == "Sunday") & (OFResult$checked[i] == FALSE)) {
    OFResult[i, 2] <- TRUE
  }
}

OFResult$date <- as.character(OFResult$date)
# change type of date to string so that rbind can be used in the loop below

# second loop runs cases (2), (4), (6)
for (i in 1:nrow(OFResult)) {
  if ((weekdays(as.Date(OFResult$date[i], "%m/%d/%Y")) == "Friday") & (OFResult$checked[i] == FALSE)) {
    if ((weekdays(as.Date(OFResult$date[i + 1], "%m/%d/%Y")) == "Saturday") & (OFResult$checked[i + 1] == FALSE)) {
      # case (2)
      j <- i + 1
      OFResult <- rbind(OFResult, c(format(as.Date(OFResult$date[i], "%m/%d/%Y") + 2, "%m/%d/%Y"), round(mean(OFResult[i:j, 2])), TRUE))
      OFResult[i:j, 3] <- TRUE
    } else {
      # case (4)
      OFResult <- rbind(OFResult, c(format(as.Date(OFResult$date[i], OFResult[i, 2], "%m/%d/%Y") + 2, "%m/%d/%Y"), TRUE))
      OFResult[i, 2] <- TRUE
    }
  }
  # case (6)
  if ((weekdays(as.Date(OFResult$date[i], "%m/%d/%Y")) == "Saturday") & (OFResult$checked[i] == FALSE)) {
    OFResult <- rbind(OFResult, c(format(as.Date(OFResult$date[i], OFResult[i, 2], "%m/%d/%Y") + 1, "%m/%d/%Y"), TRUE))
    OFResult[i, 3] <- TRUE
  }
}

OFResult$checked <- NULL
