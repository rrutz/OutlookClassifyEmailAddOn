require(nnet)

numberPrediectionToWord <- function(pred)
{
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

naiveBayesModelPredict <- function(x, parameters)
{
  x <- as.character(x)
  words <- as.data.frame(strsplit(x, split = " ")[[1]] )
  colnames(words) <- c("word")
  words$word <- words$word[!is.na(words$word)]
  words <- as.data.frame(words[words$word!="",  ])
  colnames(words) <- c("word")
  probs <- merge(x=words,y=parameters, by = "word" )[, c(3:6)]
  pred <- which.max(sapply(probs, FUN=prod))
  return(numberPrediectionToWord(pred))
}

LogisticPredict <- function(x, model)
{
  x <- as.character(x)

  objects <- c()
  for( object in x)
  {
    for( i in trimws(strsplit(object, split = ";")[[1]] ))
      objects <- c(objects,  i)
  }
  
  coeffNames <-  colnames(summary(model)[["coefficients"]])[2:(length(colnames(summary(model)[["coefficients"]]))) ]
  x <- rep(0, length(coeffNames) )
  for( object in objects)
  {
    x[coeffNames == gsub(" ", ".",gsub("[^A-Za-z///' ]","." , object ,ignore.case = TRUE))] <- 1
  }
  x <- t(data.frame(x))
  colnames(x) <- coeffNames
  p <- as.character(predict(model,x))
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


predictX <- function()
{
  df.emails <- as.data.frame(read.csv(file = signalFile, header=FALSE))
  df.emails[is.na(df.emails)] <- "None"
  colnames(df.emails) <- c( "from", "to", "cc", "subject", "body")
  
  nb.body <- unlist(lapply( df.emails$body, FUN=naiveBayesModelPredict, parameters=bodyNR), use.names=TRUE )
  nb.subject <- unlist(lapply( df.emails$subject, FUN=naiveBayesModelPredict, parameters=subjectNR), use.names=TRUE )
  lr.from <-  unlist(lapply( df.emails$from, FUN=LogisticPredict, model=fromLR), use.names=TRUE )
  lr.cc <-  unlist(lapply( df.emails$cc, FUN=LogisticPredict, model=ccLR), use.names=TRUE )
  lr.to <-  LogisticPredict(df.emails$to, model=toLR) 
 
  #X <-  matrix(  unlist( lapply( nb.body, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE) 
  #X <- cbind(X, matrix(  unlist( lapply( nb.subject, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
  ##X <- cbind(X, matrix(  unlist( lapply( lr.from, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
  #X <- cbind(X, matrix(  unlist( lapply( lr.cc, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
  #X <- cbind(X, matrix(  unlist( lapply( lr.to, FUN=dummyCode), use.names=TRUE ), ncol=4, nrow = length(lr.from), byrow=TRUE)  )
  #X <- as.data.frame(X)
  #coeffNames <-  colnames(summary(combinedLR)$coefficients)[2:(length(colnames(summary(combinedLR)$coefficients))) ]
  #colnames(X) <-  coeffNames
  #p <- as.character(predict(combinedLR,   X))
  weights <- combinedWeights
  pred <- weights$V2[1]*dummyCode(nb.body) + weights$V2[2]*dummyCode(nb.subject) + weights$V2[3]*dummyCode(lr.from) + weights$V2[4]*dummyCode(lr.cc) +weights$V2[5]*dummyCode(lr.to)
  pred <-which(pred ==max(pred))
  return(numberPrediectionToWord(pred))
}

#p <- predictX()
#cat(p)
#write.table(p, file="C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\Prediction2.txt", row.names = FALSE, col.names = FALSE, quote=FALSE)
#write.table("a", file="C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\killR", row.names = FALSE, col.names = FALSE, quote=FALSE)
#write.table("a", file="C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\EmailClassifier\\bin\\Debug\\signalFiles\\killR", row.names = FALSE, col.names = FALSE, quote=FALSE)

main <- function()
{
  args = commandArgs(trailingOnly=TRUE)
  signalFile <<- trimws(args[1])
  predictionFile <<- trimws(args[2])
  killRScriptFile <<- trimws(args[3])
  modelsDirectory <<- trimws(args[4])
  workingDirectory <<- trimws(args[5])


  bodyNR <<- read.csv( file = paste(modelsDirectory, "\\bodyNBpamramters.csv", sep=""), header=TRUE)
  subjectNR <<- read.csv( file = paste(modelsDirectory, "\\subjectNBpamramters.csv", sep=""), header=TRUE)
  ccLR <<- readRDS( file = paste(modelsDirectory, "\\ccLR.rda", sep=""))
  toLR <<- readRDS( file = paste(modelsDirectory, "\\toLR.rda", sep=""))
  fromLR <<- readRDS( file = paste(modelsDirectory, "\\fromLR.rda", sep=""))
  combinedWeights <<- read.table(file = paste(modelsDirectory, "\\weights.txt", sep=""))
  while( !file.exists(killRScriptFile) )
  {
    while( !file.exists(signalFile) & !file.exists(killRScriptFile)  )
    {
      Sys.sleep(0.001)
    }
    p <- predictX()
    write.table(p, file=predictionFile, row.names = FALSE, col.names = FALSE, quote=FALSE)
    while(file.exists(signalFile) & !file.exists(killRScriptFile))
    {
      Sys.sleep(0.001)
    }
  }
}
Sys.sleep(2)
main()
