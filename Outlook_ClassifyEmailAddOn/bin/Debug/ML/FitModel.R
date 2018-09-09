# get data
readInData <- function()
{
  df.emails <- as.data.frame(read.csv(dataFile), header=FALSE)
  colnames(df.emails) <- c("class", "from", "to", "cc", "subject", "body")
  df.emails$body <- tolower(df.emails$body )
  df.emails$body  <- gsub("[^a-z ]"," " , df.emails$body)
  return(df.emails)
}

naiveBayesModelFit <- function(data_in, columnName)
{
  # get all words
  words <- c()
  for ( body in data_in[,columnName])
  {
    words <- c(words, strsplit(body, split = " ")[[1]])
  }
  
  # create frequencies
  words.freq<-table(unlist(words))
  words.freq = as.data.frame(cbind(names(words.freq),as.integer(words.freq)) )
  words.freq$V2 <- as.numeric(words.freq$V2)
  
  getFreq <- function( class, df_in)
  {
    classEmails <- df.emails[ df.emails$class == class, ]
    words_tmp <- c()
    for ( body in classEmails$body)
    {
      words_tmp <- c(words_tmp, strsplit(body, split = " ")[[1]])
    }
    words.freq_tmp<-table(unlist(words_tmp))
    words.freq_tmp = as.data.frame(cbind(names(words.freq_tmp),as.integer(words.freq_tmp)) )
    words.freq_tmp$V2 <- as.numeric(words.freq_tmp$V2)
    words.freq_tmp$V2 <- words.freq_tmp$V2 / sum(words.freq_tmp$V2)
    
    words.freq <- merge(x=df_in, y=words.freq_tmp, by ="V1", all.x = TRUE )
    return(words.freq)
  }
  
  words.freq <- getFreq("Read", words.freq)
  words.freq <- getFreq("Delete", words.freq)
  words.freq <- getFreq("FollowUp", words.freq)
  words.freq <- getFreq("Ignore", words.freq)

  colnames(words.freq) <- c( "word", "totalCount", "readFreq", "deleteFreq", " followUpFreq", "ignoreFreq")
  words.freq[,3:6] <- words.freq[,3:6] +0.5
  words.freq[is.na(words.freq)]  <- 0.5
  write.csv(x = words.freq, file = paste(modelsDirectory,"\\",columnName, "NBpamramters.csv", sep=""), row.names = FALSE)
}

naiveBayesModelPredict <- function(x, columnName)
{
  par <- as.data.frame(read.csv(file = paste(modelsDirectory,"\\",columnName, "NBpamramters.csv", sep="")))
  x <- as.character(x)
  words <- as.data.frame(strsplit(x, split = " ")[[1]] )
  colnames(words) <- c("word")
  words$word <- words$word[!is.na(words$word)]
  words <- as.data.frame(words[words$word!="",  ])
  colnames(words) <- c("word")
  probs <- merge(x=words,y=par, by = "word" )[, c(3:6)]
  
  pred <- which.max(sapply(probs, FUN=prod))
  
  if(pred==1)
    p <- "Read"
  
  if(pred==2)
    p <- "Delete"
  
  if(pred==3)
    p <- "FollowUp"
  
  if(pred==4)
    p <- "Ignore"
  
  return(p)
  
}
  
  
require(nnet)


fitLogisticModel <- function(data_in, columnName)
{
  x <- data_in[,columnName]
  allx <- c()
  for( object in x )
  {
    for( i in trimws(strsplit(object, split = ";")[[1]] ))
      allx <- c(allx, i )
  }

  
  uniqueItems <- unique(allx)
  X <- matrix(0, ncol = length(uniqueItems), nrow = nrow(data_in))
  colnames(X) <- uniqueItems
  
  rowCount <- 1
  for( object in x )
  {
    row <- rep(0, length(uniqueItems))
    for( i in trimws(strsplit(object, split = ";")[[1]] ))
    {
      row[i == uniqueItems] <- 1
    }
    X[rowCount,  ] <- row
    rowCount <- rowCount  +1 
  }
  
  
  colnames(X) <- uniqueItems
  y <- data_in$class
  data <- cbind(y, data.frame(X))
  model <- multinom( formula =  y ~ . ,data = data) 
  
  saveRDS( model, file = paste(modelsDirectory,"\\",columnName, "LR.rda", sep=""))

}



LogisticPredict <- function(x, columnName)
{
  model <- readRDS( file = paste(modelsDirectory,"\\",columnName, "LR.rda", sep=""))
  
  objects <- c()
  for( object in x)
  {
    for( i in trimws(strsplit(object, split = ";")[[1]] ))
        objects <- c(objects,  i)
  }


  coeffNames <-  colnames(summary(model)$coefficients)[2:(length(colnames(summary(model)$coefficients))) ]
  x <- rep(0, length(coeffNames) )
  for( object in objects)
  {
    x[coeffNames == gsub(" ", ".",gsub("[^A-Za-z///' ]","." , object ,ignore.case = TRUE))] <- 1
  }
  x <- t(data.frame(x))
  colnames(x) <- coeffNames
  
  p <- as.character(predict(model,    x))
  return(p)
}

dummyCode <- function( class)
{
  d <- rep(0, 4)
  if( class == "Read" )
    d[1] <- 1
  if( class == "Delete" )
    d[2] <- 1
  if( class == "FollowUp" )
    d[3] <- 1
  if( class == "Ignore" )
    d[4] <- 1
  return(d)
}

linearCombo <- function( df.emails )
{
  nb.body <- unlist(lapply( df.emails$body, FUN=naiveBayesModelPredict, columnName="body"), use.names=TRUE )
  nb.subject <- unlist(lapply( df.emails$subject, FUN=naiveBayesModelPredict, columnName="subject"), use.names=TRUE )
  lr.from <-  unlist(lapply( df.emails$from, FUN=LogisticPredict, columnName="from"), use.names=TRUE )
  lr.cc <-  unlist(lapply( df.emails$cc, FUN=LogisticPredict, columnName="cc"), use.names=TRUE )
  lr.to <-  unlist(lapply( df.emails$to, FUN=LogisticPredict, columnName="to"), use.names=TRUE )
  
  
  #X <-  matrix(  unlist( lapply( nb.body, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE) 
  #X <- cbind(X, matrix(  unlist( lapply( nb.subject, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
  #X <- cbind(X, matrix(  unlist( lapply( lr.from, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
 # X <- cbind(X, matrix(  unlist( lapply( lr.cc, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
 # X <- cbind(X, matrix(  unlist( lapply( lr.to, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
  
  y <- df.emails$class
 # data <- cbind(y, data.frame(X))
 # model <- multinom( formula =  y ~ . ,data = data) 
  
  chooseWeights <- function(x)
  {
    a <- x[1]
    b <- x[2]
    c <- x[3]
    d <- x[4]
    e <- x[5]
    value <- -sum(a*sum(nb.body == y ) + b*sum(nb.subject == y ) + c*sum(lr.from == y ) + d*sum(lr.cc == y ) + e*sum(lr.to == y )) 
    return(value)
  }
  #constrOptim(theta =c(.2,.2,.2,.2,.2), f =chooseWeights  )
  #optim(par=c(.2,.2,.2,.2,.2), fn=chooseWeights, lower = c(0,0,0,0,0), upper = c(1,1,1,1,1), method="L-BFGS-B")


  
  write.table(as.data.frame( matrix(c("a", .05, 'b', 0.05, 'c', .2, 'd',.35,'e', .35),ncol=2, byrow = TRUE )), file =paste(modelsDirectory,"\\weights.txt", sep=""), row.names = FALSE, col.names = FALSE )
  
  
  #  coeffNames <-  colnames(summary(combinedLR)$coefficients)[2:(length(colnames(summary(combinedLR)$coefficients))) ]
  
 # colnames(X) <-  coeffNames
  
#  as.character(predict(combinedLR,   X)) == y
#  as.character(predict(model,    x))
 # saveRDS( model, file ="C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\finalLR.rda")
}

main <- function()
{
  args = commandArgs(trailingOnly=TRUE)
  modelsDirectory <<- trimws(args[1])
  dataFile <<- trimws(args[2])
   
  df.emails <<- readInData()
  df.emails <<- df.emails[!apply(df.emails == "", 1, all),]
  naiveBayesModelFit(df.emails, "body")
  naiveBayesModelFit(df.emails, "subject")
  fitLogisticModel(df.emails, "from")
  fitLogisticModel(df.emails, "cc")
  fitLogisticModel(df.emails, "to")
  linearCombo(df.emails)
}

main()
