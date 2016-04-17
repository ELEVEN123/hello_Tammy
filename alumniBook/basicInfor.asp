<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title>基本信息</title>
</head>
<body>
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

dim rs,sql,ID
ID=request.querystring("ID")
set rs=server.CreateObject("adodb.recordset")
sql="select * from 信管班成员表 where 学号='"&ID&"'"
rs.open sql,conn,1,2
response.write"<table width='290px' height='200px'><tr><td>姓名：</td><td>"&rs("姓名")&"</td></tr><tr><td>性别：</td><td>"&rs("性别")&"</td></tr><tr><td>生日：</td><td>"&rs("生日")&"</td></tr><tr><td>QQ号：</td><td>"&rs("QQ号")&"</td></tr><tr><td>手机号码：</td><td>"&rs("手机号码")&"</td></tr><tr><td>最喜欢的颜色：</td><td>"&rs("最喜欢的颜色")&"</td></tr><tr><td>个人座右铭：</td><td>"&rs("个人座右铭")&"</td></tr></table>"
%>
</body>
</html>

