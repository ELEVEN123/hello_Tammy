<!--#Include File="../messageBoard/function.asp"-->

<html>
<head>
<title>日记插入</title>
</head>
</html>
<%
'连接数据库
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

dim rs,sql,signature
signature=myDangerEncode(request.form("signature"))
set rs=server.createobject("adodb.recordset")
sql="select * from 日记表"
rs.open sql,conn,1,2
rs.addnew
rs("日记人")=session("userID")
rs("日记时间")=date()
rs("日记内容")=signature
rs.update
rs.close
set rs=nothing
response.redirect"../homepageFrameset/personalPage.asp"

%>