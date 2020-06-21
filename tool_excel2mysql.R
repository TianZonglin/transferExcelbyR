# 初始化一次即可
#install.packages("RMySQL")
#install.packages("stringr")
#install.packages("readxl")
#install.packages("readr")

library("RMySQL")   # 建立数据库连接
library("stringr")  # 引入字符串处理包
library("readxl")   # 引入excel操作包
library("readr")   # 引入csv读写

################ 准备工作

# 设置工作目录
setwd("C:\\Users\\zonglin\\OneDrive - Universiteit Utrecht\\Desktop\\ecProject\\")
# 建立Mysql连接 
conn = dbConnect(MySQL(), user = 'root', password = 'root', dbname = 'ecproject',host = 'localhost')
# 清空日志
if(file.exists("processRecord.csv")){
  file.remove("processRecord.csv")
}


# 列出所有表（测试连接）
dbListTables(conn)

 
################### 封装，读Excel，规则解析，构造sql并插入



transExcel2MysqlDB <- function(fpath,allFiles,stratmark = 1) {
  
  
  #"\\io_Input_Excel_Folder\\2016年1-10月\\re2016-03.xls"
  lp = fpath
  fpath = str_c("io_Input_Excel_Folder\\",fpath,sep="")
  edata <- readxl::read_excel(fpath) 
  edata <- edata[30:35,]  ######## 只实验6条记录
  #View(edata)
  
  # 探测数据起始行终止行
  i = 1
  iStartRow = iEndRow = 0
  sucs = errr = 0
  flag = TRUE
  while(TRUE){
    if(!is.na(as.numeric(edata[,1][i,]))){
      if(flag){
        iStartRow = i
        flag = FALSE
      }
    } 
    if(is.na(as.numeric(edata[,1][i,])) && !flag){
      iEndRow = i-1
      print(i)
      break
    } 
    i = i + 1
  }

  
  if(allFiles==0){
    dbSendQuery(conn,'SET NAMES utf8')
    dbSendQuery(conn, "SET FOREIGN_KEY_CHECKS=0;")
    dbSendQuery(conn, "DROP TABLE IF EXISTS `tb_from_excel`;")
    createSql = NULL
    for(ic in 1:(length(colnames(edata))+1)){
      if(ic == 1){
        createSql = paste(createSql,"`ID` bigint(20) unsigned NOT NULL AUTO_INCREMENT,PRIMARY KEY (`ID`)",",",sep="")
      }else if(ic==(length(colnames(edata))+1)){
        createSql = paste(createSql,"`",colnames(edata)[ic-1],"` varchar(255) DEFAULT NULL",sep="")
      }else{
        createSql = paste(createSql,"`",colnames(edata)[ic-1],"` varchar(255) DEFAULT NULL",",",sep="")
      }
      
    }
    createSql = paste("CREATE TABLE `tb_from_excel` (",createSql,") ENGINE=MyISAM DEFAULT CHARSET=utf8;",sep="")  
    print(createSql)
    dbSendQuery(conn, createSql)
    dbSendQuery(conn,"ALTER TABLE tb_from_excel AUTO_INCREMENT=1;")
    

    
  }
  allFiles = allFiles +1
  
  res <- dbSendQuery(conn,"select COLUMN_NAME from information_schema.COLUMNS where table_name = 'tb_from_excel'")
  preNames <- data.frame(dbFetch(res))[,1]
  dbClearResult(res)  
  preString = NULL
  for(i in 2:length(preNames)){
    if(i ==length(preNames)){
      preString = paste(preString,paste("`",preNames[i],"`",sep=""),sep="")
      break
    }
    preString = paste(preString,paste("`",preNames[i],"`",sep=""),",",sep="")
  }
  
  colNum = length(as.list(edata[1,]))
  if(0){
  }else{
    endSqlArr = array()
    endSqlCounter= 1;
    for (i in iStartRow:iEndRow){
      
      endString =NULL
      for(j in stratmark:colNum){
        tmpd = str_replace_all(unlist(as.list(edata[i,][j])),"'","’")
        if(j==colNum){
          endString <- paste(endString,"'",tmpd,"'",sep="")
          break
        }
        else{ 
          endString <- paste(endString,"'",gsub("[\r\n]", " ", tmpd),"'",",",sep="")
        }
        
      }
      endSqlArr[endSqlCounter] = endString
      endSqlCounter = endSqlCounter +1
      print(paste("length(endSqlArr) = ",length(endSqlArr)))
      
    }  
    for(index in 1:length(endSqlArr)){
      finalSqlString = paste("INSERT INTO tb_from_excel(",preString,") values (",endSqlArr[index],")",sep="")
      ac = NULL
      try(ac <- dbSendQuery(conn,finalSqlString), silent=TRUE)
      if(is.null(ac)){
        ac = "ErrInfo:"
        errr = errr + 1
        write(paste(ac,finalSqlString,sep="      "),"processRecord.csv",append = TRUE)
      }else{
        ac = "Accepted"
        sucs = sucs + 1
      }
      
    }
    write(paste("Summary:      ","Read success > ",lp,", ",iStartRow," : ",iEndRow,sep=""),"processRecord.csv",append = TRUE)
    write(paste("Summary:      ","Excuted ",length(endSqlArr)," SQLs",sep=""),"processRecord.csv",append = TRUE)
    write(paste("Summary:      ",sucs," Success, ",errr," Failed, ",(sucs/length(endSqlArr))*100,"% Accepted",sep=""),"processRecord.csv",append = TRUE)
    write(paste("Summary:      ","-----------------------------------",sep=""),"processRecord.csv",append = TRUE)    
    write(paste("Summary:      ",sep=""),"processRecord.csv",append = TRUE)    
    
  }
  c(sucs,errr,length(endSqlArr))
} 

################## 最上层大循环，文件读取 #############

# 开始操作读取文件
# 列出全部年份文件夹
nameAllFolders = list.files("io_Input_Excel_Folder")   
rst = c(0,0,0)
canOpen<-array()
index = 0

# 遍历每个文件夹（构造访问路径，进入，列出xls，构造访问文件路径，访问）
cnt = 0
for( folder in nameAllFolders){
  pathFolder = paste("io_Input_Excel_Folder\\",folder,"\\", sep = "")
  nameAllExcels = list.files(pathFolder)   
  for( excel in nameAllExcels){
    if(str_ends(excel, ".xls")||str_ends(excel, ".xlsx")){
      pathExcel = paste(pathFolder,excel, sep = "")
      print(pathExcel)
      content <- NULL
      try(content <- readxl::read_excel(pathExcel), silent = TRUE)
      if(is.null(content)){
        print(str_c(folder,"\\",excel,sep=""))
        canOpen[index] = str_c(folder,"\\",excel,sep="")
        index = index + 1
      }else{
        ######### 正常读取、解析 ########
        tmpPath = str_c(folder,"\\",excel,sep="")
        tmp = transExcel2MysqlDB(tmpPath,cnt)
        cnt = cnt+1
        rst = c(rst[1]+tmp[1],rst[2]+tmp[2],rst[3]+tmp[3])
      }
    }
  }
}
write(paste("ErrInfo:      ",sep=""),"processRecord.csv",append = TRUE)
for(i in 1:length(canOpen)){
  write(paste("ErrInfo:      Unread files > ",canOpen[i], sep=""),"processRecord.csv",append = TRUE)
}
write(paste("Finally:      ",sep=""),"processRecord.csv",append = TRUE)
write(paste("Finally:      ",sep=""),"processRecord.csv",append = TRUE)
write(paste("Finally:      ","The whole proccess excuted ",rst[3]," SQLs",sep=""),"processRecord.csv",append = TRUE)
write(paste("Finally:      (All together) ",rst[1]," Success, ",rst[2]," Failed, ",(rst[1]/rst[3])*100,"% Accepted",sep=""),"processRecord.csv",append = TRUE)
write(paste("Finally:      ","Transfer data from Excel to Mysql..",sep=""),"processRecord.csv",append = TRUE)
write(paste("Finally:      Finished...",sep=""),"processRecord.csv",append = TRUE)


# 收尾
dbDisconnect(conn) 


