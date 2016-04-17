


<!--#include file="upload_5xsoft.inc"--> 
<html>
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>文件上传执行</title> 
</head> 
<body>
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select * from 信管班成员表 where 学号='"&session("userID")&"'"
rs.open sql,conn,3,2
   
dim upload,file,formPath 
set upload=new upload_5xSoft
formPath=upload.form("filepath") 
if right(formPath,1)<>"/" then formPath=formPath&"/" 

for each formName in upload.file  
set file=upload.file(formName) 'file应该是fileInfo类型


 
if file.filesize<1 then 
response.write "<font size=2>请先选择你要上传的文件 <a href='photoSelect.asp'>重新上传</a></font>" 
response.end 
end if 

if file.FileSize>0 then  
file.SaveAs Server.mappath("photos\"&file.FileName)
end if

   rs("照片")=Server.mappath("photos\"&file.FileName)
   rs.update  
  
set file=nothing 
next 
set upload=nothing 
response.redirect "../homepageFrameset/homepageFrame.html" 
%> 
</body> 
</html> 
