#uses "CtrlADO.dll"
#uses "rdb.ctl"

main()
{  
 int iCount;
 dyn_anytype dsLine;
 langString header;
 bool colVisible;
 dyn_string dsHeader;
 string headerTxt;
 dyn_string dsSQL;
 string sql ;
 dbConnection db;
 string conn;
  string fileName;
  
 for(int i=0;i<table_top.columnCount();i++)
  {
    getValue("table_top","columnHeader",i,header);
    
    getValue("table_top","columnVisibility",i,colVisible);
    if(colVisible)
    {
      headerTxt = header[getLangIdx("en_US.utf8")];
      strreplace(headerTxt," ","_");
     // headerTxt = recode(headerTxt,"UTF8","GB2312");
      dynAppend(dsHeader,headerTxt);
    }
  } 
 
  if(dynlen(dsHeader)<1) 
  {
    DebugN("dsHeader length= 0");
    return;
  }
  fileSelector(fileName,PROJ_PATH,1,"*.xlsx",false);
  strreplace( fileName,"/","//");

  
  
  
  conn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};FIL=Excel 12.0;DriverID=1046;Readonly=False;DBQ=" + fileName;
  DebugN(conn);
  
// if(rdbOpen(db,"DSN=excelrw"))
  
  if(rdbOpen(db,conn))
  {
   showError(db);
   DebugN("OPEN ERROR");
   return;
  }

  if(isTableExists(db)) 
  { 
    if(rdbExecute(db,makeDynString("DROP TABLE demo")))
    {
      showError(db);
    }
  }
  
  sql = "CREATE TABLE demo(";
  
  for(int i=1;i<dynlen(dsHeader);i++)
  {
    sql = sql + dsHeader[i] +   " TEXT,"; 
  }
  
  sql = sql +    dsHeader[dynlen(dsHeader)] +  " TEXT"; 
  sql = sql  +  ")"; 
  
   //创建表格
 // if(rdbTransaction(db)) {   DebugN("rdbTransaction ERROR"); showError(db);} 
  if(rdbExecute(db,makeDynString(sql))){DebugN("rdbExecute ERROR");showError(db);} 
 // if(rdbCommit(db)){DebugN("rdbCommit ERROR");showError(db);}
   
  // 添加一行 
 for(int i=0;i<table_top.lineCount();i++)
 {
  dsLine = table_top.getLineN(i); 
  sql = "INSERT INTO demo(";
  for(int i=1;i<dynlen(dsHeader);i++)
  {
    sql = sql + (string)dsHeader[i]  + ","; 
  }
  
  sql = sql + (string)dsHeader[dynlen(dsHeader)] + ") " ;
  
  
  sql = sql + "VALUES(";
  
  for(int i=1;i<dynlen(dsHeader);i++)
  {
    sql = sql  + "'" + (string)dsLine[i] +"'" + ",";
  }
    sql = sql  + "'" + (string)dsLine[dynlen(dsHeader)] +"'"  + ")";
  
//  DebugN(sql); 
  dynAppend(dsSQL,sql); 
 }
  
 
 
  DebugN(dsSQL);
  
  if(rdbExecute(db,dsSQL)) 
  {   
    DebugN("rdbExecute ERR");   
    showError(db); 
  }
  else 
  { 
    DebugN("保存成功");
  }
  
  
//  if(rdbCommit(db)){  showError(db); } else { DebugN("保存成功");}
  
    if(rdbClose(db))
    {
    DebugN("未正常关闭excel文件" + fileName);   
    showError(db); 
    }
    else
    {
     DebugN(fileName + ".xlsb 文件保存成功");
    } 
    
    
 }
 
string GetTableLineValue(anytype value)
  {
    string sRet;
    if(getType(value) == ATIME_VAR || getType(value) == TIME_VAR )
    {
      sRet = "'" + (string)value + "'";
    }
    else
    {
      sRet = value;
    }   
    return sRet;

  }
  

  
bool isTableExists(dbConnection conn)
  {
    int errCnt, errNo, errNative;
    string errDescr, errSql;
    int rc;
    bool hasTable = false;
    rdbExecute(conn,makeDynString("CREATE TABLE demo"));
    rc = dbGetError(conn, errCnt, errNo, errNative, errDescr, errSql);
    
    errNative = recode(errNative,"GB2312","UTF8");
    errDescr = recode(errDescr,"GB2312","UTF8");
    errSql = recode(errSql,"GB2312","UTF8");
 
    if(strpos(errDescr,"已存在")>0)
    {
      hasTable = true;
    } 
    return hasTable; 
  }
  
  


void showError (dbConnection conn)
{
  int errCnt, errNo, errNative;
  string errDescr, errSql;
  int rc;
  errCnt = 1;
  rc = 0;
  while (errCnt > 0 && ! rc)
  {
    rc = dbGetError(conn, errCnt, errNo, errNative, errDescr, errSql); 
    errNative = recode(errNative,"GB2312","UTF8");
    errDescr = recode(errDescr,"GB2312","UTF8");
    errSql = recode(errSql,"GB2312","UTF8"); 
  if (!rc)
  {
    DebugN("Errornumber : ", errNo);
    DebugN("BaseError : ", errNative);
    DebugN("Description : ", errDescr);
    DebugN("SQL-errortext: ",errSql);
  }
  else
    DebugN("dbGetError failed, rc = ", rc);
    errCnt--;
  }

}

