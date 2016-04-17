<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 
%>

		
<html>
<head>
	<title>删除留言</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
	
	<%
	
	dim txtID,rs,sql
	txtID=Request.querystring("txtID")
	set rs=Server.CreateObject("ADODB.recordset")
	sql="Select 留言人 From 留言板记录表 where ID=" &txtID
	rs.open sql,conn,1,2
	
	If rs("留言人")=session("userID") Then
		Dim strSql
		strSql="Delete From 留言板记录表 where ID=" &txtID 
		conn.Execute(strSql)
		Response.Redirect("index.asp")
	else response.write"对不起，你不能删除别人的留言!"
	End If
	%>
