![](https://cdn.jsdelivr.net/gh/TianZonglin/tuchuang/img/20200621141106.png)


# transferExcelbyR

### suitable scenarios and requirements

- large number of Excels stored in folders
- ALL Excels MUST have the same data (column) format 
- only with 1 sheet in each Excels
- Excels must located in the third level folder, example:

```
ecProject\io_Input_Excel_Folder\simples\ORGDATA.XLS
# workDir -> inputFolder(1st) -> simples(2nd) -> realExcel(3rd)
```

### feature

- automatically create table using excel's column name
- automatically detect the region (start/end) of Excels
- detailed logging info
- transfer Excels in folders or whatever
- combine *multiple* excel files into *one* DB table (.sql)

### Usage

#### common tool is

> tool_excel2mysql.R

Recommend to use R Studio to run it.
 
#### softwares and dev environment

![](https://cdn.jsdelivr.net/gh/TianZonglin/tuchuang/img/20200621115758.png)


You can find them all on the Internet.

#### install packages we need

```
# just run them once, near line 8
#install.packages("RMySQL")
#install.packages("stringr")
#install.packages("readxl")
#install.packages("readr")
```

#### change workdir

```
# near line 14
setwd("C:\\Users\\zonglin\\OneDrive - Universiteit Utrecht\\Desktop\\ecProject\\")
```

#### change mysql connect configuration

```
# default database name: test
# near line 17
conn = dbConnect(MySQL(), user = 'root', password = 'root', dbname = 'test',host = 'localhost')

# defalut table namme: tb_from_excel
# use editor's find/replace function to replace it all.
```

#### select suitable start position (column)

```
# default start columns number: 1
# near line 222
tmp = transExcel2MysqlDB(tmpPath, cnt, startmark = 2)
```

#### test part of data

If you have a huge number of Excels and you just wanna test this code or catch the debug information of Excels (can open or not) with the `errinfo with finally` in logs, you can modify the row number below. Then it just takes limited row data with every Excels. 

```
# near line 85
edata <- edata[30:35,] 
```

### Logs (processRecord.csv)

![](https://cdn.jsdelivr.net/gh/TianZonglin/tuchuang/img/20200621113612.png)

#### errinfo with summary

This is the record of failed insert-sqls. If you use folders to contain more than one Excels, then every excel could output a part of `errinfo with summary`. Using this cache info we can find the wrong sql items with the help of Navicat, which could automatically valid the wrong position easily.

![](https://cdn.jsdelivr.net/gh/TianZonglin/tuchuang/img/20200621114705.png)

Then you can modify the code of `tool_excel2mysql` to fix it or just give me feedback.

#### errinfo with finally

This is the global information with the unreadable Excels and final summaries. If one excel appears here, then you need to check this file manually to find what's the real problem it has. Sometimes it could rerun well after resave (open it then save it) these Excels by your hands. 

Basically the tool could transfer data from (my) xls, xlsx files to mysql smoothly with almost 100% success rate. (that screenshot was a demo to show errinfo)

---

[中文说明](https://www.cz5h.com/article/528.html)

---

**ENJOY...**
